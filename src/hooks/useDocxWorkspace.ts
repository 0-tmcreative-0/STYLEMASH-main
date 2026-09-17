import { useCallback, useMemo, useReducer, useRef } from 'react'
import type {
  FormattingSignature,
  ListFormat,
  ParsedDocx,
  StyleEntity,
  UserStyleKind,
  UserStyleRecord,
} from '../types/ooxml'
import { downloadBlob } from '../lib/download'
import {
  bulkMergeVariantsIntoMatchingReferenceStyles,
  findVariantIdsMatchingReferenceStyleNames,
} from '../lib/ooxml/bulkMergeMatchedStyles'
import { buildContentMergedDocx, type ContentMergeOptions } from '../lib/ooxml/contentMerge'
import { addDefaultStyles } from '../lib/ooxml/defaultStyles'
import { mergeParagraphStyle, mergeStyles, removeStyleById } from '../lib/ooxml/mergeStyles'
import { parseDocx } from '../lib/ooxml/parseDocx'
import {
  materializeReferenceDocStyles,
  reconcileUserStylesOnReferenceDocRemoval,
} from '../lib/ooxml/referenceDocStyles'
import { serializeDocx, serializePart } from '../lib/ooxml/serializeDocx'
import { buildStyleReport, collectRunRefsForVariantIds, findVariantById } from '../lib/ooxml/styleReport'
import { applyXmlFragmentToRunRefs } from '../lib/ooxml/xmlFragmentEdit'

/** A pre-mutation snapshot of everything an undoable action (a merge, an
 * "Edit XML" apply, or a full "Clear list") touches, captured right before
 * the mutating call - documentXml/stylesXml/numberingXml are mutated in
 * place (see the reducer's top-of-file note), so the only way to snapshot
 * them is to serialize to text and re-parse on restore, rather than keeping
 * a second live reference to the same, about-to-be-mutated Elements. */
interface UndoSnapshot {
  documentXml: string
  stylesXml: string
  numberingXml: string | null
  userStyles: UserStyleRecord[]
}

function snapshotForUndo(parsedDocx: ParsedDocx, userStyles: UserStyleRecord[]): UndoSnapshot {
  return {
    documentXml: serializePart(parsedDocx.documentXml),
    stylesXml: serializePart(parsedDocx.stylesXml),
    numberingXml: parsedDocx.numberingXml ? serializePart(parsedDocx.numberingXml) : null,
    userStyles,
  }
}

function restoreSnapshot(parsedDocx: ParsedDocx, snapshot: UndoSnapshot): ParsedDocx {
  const parser = new DOMParser()
  return {
    ...parsedDocx,
    documentXml: parser.parseFromString(snapshot.documentXml, 'application/xml'),
    stylesXml: parser.parseFromString(snapshot.stylesXml, 'application/xml'),
    numberingXml: snapshot.numberingXml
      ? parser.parseFromString(snapshot.numberingXml, 'application/xml')
      : null,
  }
}

export interface ReferenceDocState {
  status: 'empty' | 'loading' | 'loaded' | 'error'
  errorMessage: string | null
  parsedDocx: ParsedDocx | null
}

const initialReferenceDocState: ReferenceDocState = {
  status: 'empty',
  errorMessage: null,
  parsedDocx: null,
}

export interface WorkspaceState {
  status: 'empty' | 'loading' | 'loaded' | 'error'
  errorMessage: string | null
  parsedDocx: ParsedDocx | null
  styleReport: StyleEntity[]
  userStyles: UserStyleRecord[]
  /** Selection operates at the variant level (see StyleEntityVariant) so a
   * style-derived instance and a direct-override instance of the same
   * visual look can be selected/merged independently. */
  selectedVariantIds: Set<string>
  /** The single User-Created style currently picked as a merge target (a
   * click on its row in UserStylesPanel, distinct from that row's "Edit"
   * button) - lets selectedVariantIds be folded into it directly via
   * mergeSelectedIntoTarget, without going through MergeDialog. */
  selectedTargetStyleId: string | null
  activeEditVariantId: string | null
  mergeDialogOpen: boolean
  /** Set when the merge dialog was opened to edit/extend an existing
   * UserStyleRecord (via UserStylesPanel's "Edit" button) rather than to
   * create a brand-new style from a fresh Style Report selection. */
  mergeDialogReuseStyleId: string | null
  mergeError: string | null
  xmlEditorError: string | null
  isSaving: boolean
  /** Document B: a second, separate document attached purely as a text
   * style reference (see referenceDocStyles.ts/contentMerge.ts). Entirely
   * independent of the merge/XML-editor modal state above. */
  referenceDoc: ReferenceDocState
  contentMergeDialogOpen: boolean
  isMergingContent: boolean
  contentMergeError: string | null
  /** Error from the last bulkMergeMatchedToReference - surfaced inline near
   * the Style Report's bulk-merge control rather than in a modal, since
   * that action has no modal of its own. */
  bulkMergeError: string | null
  /** One entry per undoable action (a merge, an "Edit XML" apply, or
   * "Clear list"), oldest first - UNDO_APPLIED pops the last one and
   * restores it. Deliberately doesn't cover style-only actions with no
   * document effect until merged into (DEFAULT_STYLES_ADDED,
   * attaching/removing Document B) - see the action creators below. */
  undoStack: UndoSnapshot[]
}

const initialState: WorkspaceState = {
  status: 'empty',
  errorMessage: null,
  parsedDocx: null,
  styleReport: [],
  userStyles: [],
  selectedVariantIds: new Set(),
  selectedTargetStyleId: null,
  activeEditVariantId: null,
  mergeDialogOpen: false,
  mergeDialogReuseStyleId: null,
  mergeError: null,
  xmlEditorError: null,
  isSaving: false,
  referenceDoc: initialReferenceDocState,
  contentMergeDialogOpen: false,
  isMergingContent: false,
  contentMergeError: null,
  bulkMergeError: null,
  undoStack: [],
}

/** Which inline error slot a failed action reports into - each of the three
 * mutating flows surfaces its error in a different part of the UI. */
type FailureScope = 'merge' | 'xmlEditor' | 'bulkMerge'

type Action =
  | { type: 'LOADING_STARTED' }
  | { type: 'FILE_LOADED'; parsedDocx: ParsedDocx; styleReport: StyleEntity[] }
  | { type: 'PARSE_ERROR'; message: string }
  | { type: 'TOGGLE_SELECT_VARIANT'; variantId: string }
  | { type: 'CLEAR_SELECTION' }
  | { type: 'OPEN_MERGE_DIALOG'; reuseExistingStyleId?: string }
  | { type: 'CLOSE_MODALS' }
  | {
      type: 'MERGE_APPLIED'
      parsedDocx: ParsedDocx
      styleReport: StyleEntity[]
      userStyles: UserStyleRecord[]
      undoSnapshot: UndoSnapshot
    }
  | { type: 'ACTION_FAILED'; scope: FailureScope; message: string }
  | { type: 'OPEN_XML_EDITOR'; variantId: string }
  | {
      type: 'XML_EDIT_APPLIED'
      parsedDocx: ParsedDocx
      styleReport: StyleEntity[]
      undoSnapshot: UndoSnapshot
    }
  | { type: 'SAVING_STARTED' }
  | { type: 'SAVING_FINISHED' }
  | { type: 'REFERENCE_DOC_LOADING_STARTED' }
  | {
      type: 'REFERENCE_DOC_LOADED'
      referenceParsedDocx: ParsedDocx
      parsedDocx: ParsedDocx
      styleReport: StyleEntity[]
      userStyles: UserStyleRecord[]
    }
  | { type: 'REFERENCE_DOC_LOAD_ERROR'; message: string }
  /** `parsedDocx` is null only in the degenerate "Document B attached with
   * no Document A loaded" case, which the UI can't actually reach. */
  | { type: 'REFERENCE_DOC_REMOVED'; parsedDocx: ParsedDocx | null; userStyles: UserStyleRecord[] }
  | { type: 'OPEN_CONTENT_MERGE_DIALOG' }
  | { type: 'CONTENT_MERGE_STARTED' }
  | { type: 'CONTENT_MERGE_FINISHED' }
  | { type: 'CONTENT_MERGE_ERROR'; message: string }
  | { type: 'SELECT_VARIANTS'; variantIds: string[] }
  | {
      type: 'BULK_MERGE_APPLIED'
      parsedDocx: ParsedDocx
      styleReport: StyleEntity[]
      undoSnapshot: UndoSnapshot
    }
  | { type: 'RESET' }
  | { type: 'TOGGLE_SELECT_TARGET_STYLE'; styleId: string }
  | {
      type: 'DEFAULT_STYLES_ADDED'
      parsedDocx: ParsedDocx
      styleReport: StyleEntity[]
      userStyles: UserStyleRecord[]
    }
  | {
      type: 'USER_STYLES_CLEARED'
      parsedDocx: ParsedDocx
      styleReport: StyleEntity[]
      undoSnapshot: UndoSnapshot
    }
  | {
      type: 'UNDO_APPLIED'
      parsedDocx: ParsedDocx
      styleReport: StyleEntity[]
      userStyles: UserStyleRecord[]
    }

/** Note on the reducer's relationship to immutability: `parsedDocx`'s inner
 * XMLDocuments (documentXml/stylesXml/numberingXml) are mutated in place by
 * mergeStyles() and applyXmlFragmentToRunRefs() rather than treated as
 * immutable data - a deliberate escape hatch. Re-parsing the whole XML tree
 * on every micro-edit would be wasteful and would invalidate the very live
 * Element references (RunRef) the next operation needs.
 *
 * Crucially, though, none of that mutation happens *here*. React may invoke
 * a reducer more than once for a single dispatch (StrictMode does so on
 * every dispatch in development, to surface exactly this class of bug), so
 * a reducer that called mergeStyles() itself would apply the merge twice -
 * creating a second, orphaned <w:style> (and, for a list style, a second
 * orphaned numbering definition) in the saved document every single time.
 * Every mutating call therefore lives in the action creators below, which
 * run exactly once per user gesture; the reducer only ever stores the
 * already-computed result. It still returns a *new* `parsedDocx` wrapper
 * object on each of those actions purely so React's shallow comparison
 * re-renders, even though the Document instances inside it are the same
 * mutated objects. */
function reducer(state: WorkspaceState, action: Action): WorkspaceState {
  switch (action.type) {
    case 'LOADING_STARTED':
      return { ...initialState, status: 'loading' }

    case 'FILE_LOADED':
      return {
        ...initialState,
        status: 'loaded',
        parsedDocx: action.parsedDocx,
        styleReport: action.styleReport,
      }

    case 'PARSE_ERROR':
      return { ...initialState, status: 'error', errorMessage: action.message }

    case 'TOGGLE_SELECT_VARIANT': {
      const next = new Set(state.selectedVariantIds)
      if (next.has(action.variantId)) next.delete(action.variantId)
      else next.add(action.variantId)
      return { ...state, selectedVariantIds: next }
    }

    case 'CLEAR_SELECTION':
      return { ...state, selectedVariantIds: new Set() }

    case 'OPEN_MERGE_DIALOG':
      return {
        ...state,
        mergeDialogOpen: true,
        mergeError: null,
        mergeDialogReuseStyleId: action.reuseExistingStyleId ?? null,
      }

    case 'CLOSE_MODALS':
      return {
        ...state,
        mergeDialogOpen: false,
        mergeDialogReuseStyleId: null,
        activeEditVariantId: null,
        mergeError: null,
        xmlEditorError: null,
        contentMergeDialogOpen: false,
        contentMergeError: null,
      }

    // Covers both merge entry points (the dialog's "Create style"/"Save
    // changes", and UserStylesPanel's "Merge N selected here"); closing the
    // dialog is a no-op for the latter, which never opened one.
    case 'MERGE_APPLIED':
      return {
        ...state,
        parsedDocx: action.parsedDocx,
        styleReport: action.styleReport,
        userStyles: action.userStyles,
        selectedVariantIds: new Set(),
        selectedTargetStyleId: null,
        mergeDialogOpen: false,
        mergeDialogReuseStyleId: null,
        mergeError: null,
        undoStack: [...state.undoStack, action.undoSnapshot],
      }

    case 'ACTION_FAILED':
      if (action.scope === 'merge') return { ...state, mergeError: action.message }
      if (action.scope === 'xmlEditor') return { ...state, xmlEditorError: action.message }
      return { ...state, bulkMergeError: action.message }

    case 'OPEN_XML_EDITOR':
      return { ...state, activeEditVariantId: action.variantId, xmlEditorError: null }

    case 'XML_EDIT_APPLIED':
      return {
        ...state,
        parsedDocx: action.parsedDocx,
        styleReport: action.styleReport,
        activeEditVariantId: null,
        xmlEditorError: null,
        undoStack: [...state.undoStack, action.undoSnapshot],
      }

    case 'SAVING_STARTED':
      return { ...state, isSaving: true }

    case 'SAVING_FINISHED':
      return { ...state, isSaving: false }

    case 'REFERENCE_DOC_LOADING_STARTED':
      return { ...state, referenceDoc: { status: 'loading', errorMessage: null, parsedDocx: null } }

    case 'REFERENCE_DOC_LOADED':
      return {
        ...state,
        referenceDoc: { status: 'loaded', errorMessage: null, parsedDocx: action.referenceParsedDocx },
        // Already the full replacement list (materializeReferenceDocStyles
        // folds a duplicate-named style into its existing record - same
        // styleId, redefined look - rather than appending a second one; see
        // referenceDocStyles.ts), so this assigns directly instead of
        // spreading onto state.userStyles.
        userStyles: action.userStyles,
        // parsedDocx (A) was mutated in place by materializeReferenceDocStyles.
        // styleReport DOES need recomputing (unlike a plain "new style with
        // no occurrences yet" used to be able to skip): a duplicate-name
        // collision redefines an *existing* <w:style>'s rPr/pPr in place,
        // and any run in A already merged into that style needs its Style
        // Report signature (and therefore its Document Preview look)
        // refreshed to match the newest file's definition.
        parsedDocx: action.parsedDocx,
        styleReport: action.styleReport,
      }

    case 'REFERENCE_DOC_LOAD_ERROR':
      return { ...state, referenceDoc: { status: 'error', errorMessage: action.message, parsedDocx: null } }

    case 'REFERENCE_DOC_REMOVED':
      return {
        ...state,
        parsedDocx: action.parsedDocx,
        userStyles: action.userStyles,
        referenceDoc: initialReferenceDocState,
        // styleReport unchanged: removal only drops zero-occurrence style
        // *definitions*, which by definition no run points at.
        // Clear a dangling target selection if the style it pointed to was
        // one of the just-removed zero-occurrence ones.
        selectedTargetStyleId: action.userStyles.some((r) => r.styleId === state.selectedTargetStyleId)
          ? state.selectedTargetStyleId
          : null,
      }

    case 'OPEN_CONTENT_MERGE_DIALOG':
      return { ...state, contentMergeDialogOpen: true, contentMergeError: null }

    case 'CONTENT_MERGE_STARTED':
      return { ...state, isMergingContent: true, contentMergeError: null }

    case 'CONTENT_MERGE_FINISHED':
      return { ...state, isMergingContent: false, contentMergeDialogOpen: false }

    case 'CONTENT_MERGE_ERROR':
      return { ...state, isMergingContent: false, contentMergeError: action.message }

    case 'SELECT_VARIANTS': {
      const next = new Set(state.selectedVariantIds)
      for (const id of action.variantIds) next.add(id)
      return { ...state, selectedVariantIds: next, bulkMergeError: null }
    }

    case 'BULK_MERGE_APPLIED':
      return {
        ...state,
        parsedDocx: action.parsedDocx,
        styleReport: action.styleReport,
        selectedVariantIds: new Set(),
        bulkMergeError: null,
        undoStack: [...state.undoStack, action.undoSnapshot],
      }

    case 'RESET':
      return initialState

    case 'TOGGLE_SELECT_TARGET_STYLE':
      return {
        ...state,
        selectedTargetStyleId: state.selectedTargetStyleId === action.styleId ? null : action.styleId,
      }

    case 'DEFAULT_STYLES_ADDED':
      return {
        ...state,
        parsedDocx: action.parsedDocx,
        userStyles: action.userStyles,
        // A name collision with an existing style redefines its rPr/pPr in
        // place (see addDefaultStyles) - same reasoning as
        // REFERENCE_DOC_LOADED for recomputing this: any run already merged
        // into that style needs its signature refreshed to match the newest
        // look.
        styleReport: action.styleReport,
      }

    case 'USER_STYLES_CLEARED':
      return {
        ...state,
        parsedDocx: action.parsedDocx,
        styleReport: action.styleReport,
        userStyles: [],
        selectedTargetStyleId: null,
        undoStack: [...state.undoStack, action.undoSnapshot],
      }

    case 'UNDO_APPLIED':
      return {
        ...state,
        parsedDocx: action.parsedDocx,
        styleReport: action.styleReport,
        userStyles: action.userStyles,
        undoStack: state.undoStack.slice(0, -1),
        selectedVariantIds: new Set(),
        selectedTargetStyleId: null,
        mergeError: null,
        bulkMergeError: null,
      }

    default:
      return state
  }
}

/** Owns the entire StyleMash workspace: the parsed document, the derived
 * Style Report, user-created (merged) styles, and all selection/modal UI
 * state. A single reducer (rather than several useStates) because a merge
 * is one atomic transaction that must update selection, the report, and the
 * user-styles list together.
 *
 * Every action creator that mutates the document does its work here, before
 * dispatching, and dispatches only the finished result - see the reducer's
 * own note on why that separation is load-bearing rather than stylistic. */
export function useDocxWorkspace() {
  const [state, dispatch] = useReducer(reducer, initialState)

  // Latest-state ref, so the mutating action creators below can read current
  // state without taking it as a dependency (which would rebuild every
  // callback on every keystroke-level state change). Assigned during render
  // rather than in an effect so a callback fired in the same commit as a
  // state change still sees the new state; the assignment is idempotent, so
  // StrictMode's double render is harmless.
  const stateRef = useRef(state)
  stateRef.current = state

  const loadFile = useCallback(async (file: File) => {
    dispatch({ type: 'LOADING_STARTED' })
    try {
      const parsedDocx = await parseDocx(file)
      const styleReport = buildStyleReport(parsedDocx)
      dispatch({ type: 'FILE_LOADED', parsedDocx, styleReport })
    } catch (err) {
      dispatch({
        type: 'PARSE_ERROR',
        message: err instanceof Error ? err.message : 'Failed to load this file.',
      })
    }
  }, [])

  const toggleSelectVariant = useCallback((variantId: string) => {
    dispatch({ type: 'TOGGLE_SELECT_VARIANT', variantId })
  }, [])

  const clearSelection = useCallback(() => dispatch({ type: 'CLEAR_SELECTION' }), [])
  const openMergeDialog = useCallback((reuseExistingStyleId?: string) => {
    dispatch({ type: 'OPEN_MERGE_DIALOG', reuseExistingStyleId })
  }, [])
  const closeModals = useCallback(() => dispatch({ type: 'CLOSE_MODALS' }), [])

  const confirmMerge = useCallback(
    (
      targetProps: FormattingSignature,
      name: string,
      kind: UserStyleKind,
      listFormat: ListFormat,
      reuseExistingStyleId?: string,
    ) => {
      const current = stateRef.current
      if (!current.parsedDocx) return
      // Empty selection is valid: it's how "+ New Style" and "redefine this
      // style's look" (no newly-selected variants to fold in) both work -
      // mergeStyles() just creates/redefines the style definition itself.
      const sourceRunRefs = collectRunRefsForVariantIds(current.styleReport, current.selectedVariantIds)
      const undoSnapshot = snapshotForUndo(current.parsedDocx, current.userStyles)

      try {
        const styleId =
          kind === 'paragraph'
            ? mergeParagraphStyle(
                current.parsedDocx,
                sourceRunRefs,
                targetProps,
                name,
                listFormat,
                reuseExistingStyleId,
              )
            : mergeStyles(current.parsedDocx, sourceRunRefs, targetProps, name, reuseExistingStyleId)

        const existingIndex = current.userStyles.findIndex((r) => r.styleId === styleId)
        const effectiveListFormat = kind === 'paragraph' ? listFormat : 'none'
        const record: UserStyleRecord = {
          styleId,
          name,
          targetSignature: targetProps,
          kind,
          listFormat: effectiveListFormat,
          // Manually-created lists are always single-level (createListNumId
          // only ever creates one), so this is just the ilvl-0 marker - the
          // richer multilevel preview (e.g. "1.1.") is exclusive to styles
          // materialized from Document B (see referenceDocStyles.ts).
          listPreviewText:
            effectiveListFormat === 'bullet' ? '•' : effectiveListFormat === 'decimal' ? '1.' : undefined,
          createdAt: existingIndex === -1 ? Date.now() : current.userStyles[existingIndex].createdAt,
        }
        const userStyles =
          existingIndex === -1
            ? [...current.userStyles, record]
            : current.userStyles.map((r, i) => (i === existingIndex ? record : r))

        dispatch({
          type: 'MERGE_APPLIED',
          parsedDocx: { ...current.parsedDocx },
          styleReport: buildStyleReport(current.parsedDocx),
          userStyles,
          undoSnapshot,
        })
      } catch (err) {
        dispatch({
          type: 'ACTION_FAILED',
          scope: 'merge',
          message: err instanceof Error ? err.message : 'Merge failed.',
        })
      }
    },
    [],
  )

  const openXmlEditor = useCallback((variantId: string) => {
    dispatch({ type: 'OPEN_XML_EDITOR', variantId })
  }, [])

  const applyXmlEdit = useCallback((fragmentText: string) => {
    const current = stateRef.current
    if (!current.parsedDocx || !current.activeEditVariantId) return
    const variant = findVariantById(current.styleReport, current.activeEditVariantId)
    if (!variant) return

    const undoSnapshot = snapshotForUndo(current.parsedDocx, current.userStyles)
    try {
      applyXmlFragmentToRunRefs(current.parsedDocx, variant.runRefs, fragmentText)
      dispatch({
        type: 'XML_EDIT_APPLIED',
        parsedDocx: { ...current.parsedDocx },
        styleReport: buildStyleReport(current.parsedDocx),
        undoSnapshot,
      })
    } catch (err) {
      dispatch({
        type: 'ACTION_FAILED',
        scope: 'xmlEditor',
        message: err instanceof Error ? err.message : 'Could not apply XML.',
      })
    }
  }, [])

  const save = useCallback(async () => {
    const parsedDocx = stateRef.current.parsedDocx
    if (!parsedDocx) return
    dispatch({ type: 'SAVING_STARTED' })
    try {
      const { blob, filename } = await serializeDocx(parsedDocx)
      downloadBlob(blob, filename)
    } finally {
      dispatch({ type: 'SAVING_FINISHED' })
    }
  }, [])

  const loadReferenceDoc = useCallback(async (file: File) => {
    const current = stateRef.current
    if (!current.parsedDocx) return
    dispatch({ type: 'REFERENCE_DOC_LOADING_STARTED' })
    try {
      const referenceParsedDocx = await parseDocx(file)
      const userStyles = materializeReferenceDocStyles(
        current.parsedDocx,
        referenceParsedDocx,
        current.userStyles,
      )
      dispatch({
        type: 'REFERENCE_DOC_LOADED',
        referenceParsedDocx,
        parsedDocx: { ...current.parsedDocx },
        styleReport: buildStyleReport(current.parsedDocx),
        userStyles,
      })
    } catch (err) {
      dispatch({
        type: 'REFERENCE_DOC_LOAD_ERROR',
        message: err instanceof Error ? err.message : 'Failed to load this file.',
      })
    }
  }, [])

  const removeReferenceDoc = useCallback(() => {
    const current = stateRef.current
    if (!current.parsedDocx) {
      dispatch({ type: 'REFERENCE_DOC_REMOVED', parsedDocx: null, userStyles: current.userStyles })
      return
    }
    const userStyles = reconcileUserStylesOnReferenceDocRemoval(
      current.userStyles,
      current.styleReport,
      current.parsedDocx.stylesXml,
    )
    dispatch({ type: 'REFERENCE_DOC_REMOVED', parsedDocx: { ...current.parsedDocx }, userStyles })
  }, [])

  const openContentMergeDialog = useCallback(() => dispatch({ type: 'OPEN_CONTENT_MERGE_DIALOG' }), [])

  const mergeContentIntoReferenceDoc = useCallback(async (keepOriginalFormatting: boolean) => {
    const current = stateRef.current
    if (!current.parsedDocx || !current.referenceDoc.parsedDocx) return
    dispatch({ type: 'CONTENT_MERGE_STARTED' })
    try {
      const options: ContentMergeOptions = { keepOriginalFormatting }
      const { blob, filename } = await buildContentMergedDocx(
        current.parsedDocx,
        current.referenceDoc.parsedDocx,
        options,
      )
      downloadBlob(blob, filename)
      dispatch({ type: 'CONTENT_MERGE_FINISHED' })
    } catch (err) {
      dispatch({
        type: 'CONTENT_MERGE_ERROR',
        message: err instanceof Error ? err.message : 'Could not merge content into Document B.',
      })
    }
  }, [])

  const selectVariantsMatchingReferenceStyles = useCallback(() => {
    const current = stateRef.current
    const variantIds = Array.from(
      findVariantIdsMatchingReferenceStyleNames(current.styleReport, current.userStyles),
    )
    dispatch({ type: 'SELECT_VARIANTS', variantIds })
  }, [])

  const bulkMergeMatchedToReference = useCallback(() => {
    const current = stateRef.current
    if (!current.parsedDocx) return
    const undoSnapshot = snapshotForUndo(current.parsedDocx, current.userStyles)
    try {
      bulkMergeVariantsIntoMatchingReferenceStyles(
        current.parsedDocx,
        current.styleReport,
        current.selectedVariantIds,
        current.userStyles,
      )
      dispatch({
        type: 'BULK_MERGE_APPLIED',
        parsedDocx: { ...current.parsedDocx },
        styleReport: buildStyleReport(current.parsedDocx),
        undoSnapshot,
      })
    } catch (err) {
      dispatch({
        type: 'ACTION_FAILED',
        scope: 'bulkMerge',
        message: err instanceof Error ? err.message : 'Bulk merge failed.',
      })
    }
  }, [])

  /** Returns to the landing screen in-app (no browser reload) - used by
   * StyleReportPanel's "Mash a different file" button, so the persistent
   * header/footer never flicker/remount along the way. */
  const reset = useCallback(() => dispatch({ type: 'RESET' }), [])

  const toggleSelectTargetStyle = useCallback((styleId: string) => {
    dispatch({ type: 'TOGGLE_SELECT_TARGET_STYLE', styleId })
  }, [])

  const mergeSelectedIntoTarget = useCallback(() => {
    const current = stateRef.current
    if (!current.parsedDocx || !current.selectedTargetStyleId) return
    const targetRecord = current.userStyles.find((r) => r.styleId === current.selectedTargetStyleId)
    if (!targetRecord) return

    const sourceRunRefs = collectRunRefsForVariantIds(current.styleReport, current.selectedVariantIds)
    const undoSnapshot = snapshotForUndo(current.parsedDocx, current.userStyles)
    try {
      if (targetRecord.kind === 'paragraph') {
        mergeParagraphStyle(
          current.parsedDocx,
          sourceRunRefs,
          targetRecord.targetSignature,
          targetRecord.name,
          targetRecord.listFormat,
          targetRecord.styleId,
        )
      } else {
        mergeStyles(
          current.parsedDocx,
          sourceRunRefs,
          targetRecord.targetSignature,
          targetRecord.name,
          targetRecord.styleId,
        )
      }
      dispatch({
        type: 'MERGE_APPLIED',
        parsedDocx: { ...current.parsedDocx },
        styleReport: buildStyleReport(current.parsedDocx),
        userStyles: current.userStyles,
        undoSnapshot,
      })
    } catch (err) {
      dispatch({
        type: 'ACTION_FAILED',
        scope: 'merge',
        message: err instanceof Error ? err.message : 'Merge failed.',
      })
    }
  }, [])

  const addDefaultStylesAction = useCallback(() => {
    const current = stateRef.current
    if (!current.parsedDocx) return
    const userStyles = addDefaultStyles(current.parsedDocx, current.userStyles)
    dispatch({
      type: 'DEFAULT_STYLES_ADDED',
      parsedDocx: { ...current.parsedDocx },
      styleReport: buildStyleReport(current.parsedDocx),
      userStyles,
    })
  }, [])

  const clearUserStyles = useCallback(() => {
    const current = stateRef.current
    if (!current.parsedDocx || current.userStyles.length === 0) return
    const undoSnapshot = snapshotForUndo(current.parsedDocx, current.userStyles)
    for (const record of current.userStyles) {
      removeStyleById(current.parsedDocx.stylesXml, record.styleId)
    }
    dispatch({
      type: 'USER_STYLES_CLEARED',
      parsedDocx: { ...current.parsedDocx },
      styleReport: buildStyleReport(current.parsedDocx),
      undoSnapshot,
    })
  }, [])

  const undo = useCallback(() => {
    const current = stateRef.current
    if (!current.parsedDocx || current.undoStack.length === 0) return
    const snapshot = current.undoStack[current.undoStack.length - 1]
    const parsedDocx = restoreSnapshot(current.parsedDocx, snapshot)
    dispatch({
      type: 'UNDO_APPLIED',
      parsedDocx,
      styleReport: buildStyleReport(parsedDocx),
      userStyles: snapshot.userStyles,
    })
  }, [])

  // Summary of the current selection, for MergeDialog's prefill/messaging -
  // exposed as derived totals rather than raw entities/variants so the
  // dialog stays decoupled from the report's grouping shape.
  const selection = useMemo(() => {
    let totalOccurrences = 0
    let baselineSignature: FormattingSignature | null = null
    for (const entity of state.styleReport) {
      for (const variant of entity.variants) {
        if (state.selectedVariantIds.has(variant.id)) {
          totalOccurrences += variant.occurrenceCount
          if (!baselineSignature) baselineSignature = entity.signature
        }
      }
    }
    return { totalOccurrences, baselineSignature }
  }, [state.styleReport, state.selectedVariantIds])

  const activeEditVariant = useMemo(
    () => (state.activeEditVariantId ? findVariantById(state.styleReport, state.activeEditVariantId) : null),
    [state.styleReport, state.activeEditVariantId],
  )

  return {
    state,
    selection,
    activeEditVariant,
    actions: {
      loadFile,
      toggleSelectVariant,
      clearSelection,
      openMergeDialog,
      closeModals,
      confirmMerge,
      openXmlEditor,
      applyXmlEdit,
      save,
      loadReferenceDoc,
      removeReferenceDoc,
      openContentMergeDialog,
      mergeContentIntoReferenceDoc,
      selectVariantsMatchingReferenceStyles,
      bulkMergeMatchedToReference,
      reset,
      toggleSelectTargetStyle,
      mergeSelectedIntoTarget,
      addDefaultStyles: addDefaultStylesAction,
      clearUserStyles,
      undo,
    },
  }
}
