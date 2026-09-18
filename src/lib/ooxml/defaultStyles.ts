import type { FormattingSignature, ListFormat, ParsedDocx, UserStyleKind, UserStyleRecord } from '../../types/ooxml'
import { mergeParagraphStyle, mergeStyles } from './mergeStyles'

/** Groups DEFAULT_STYLES for display in the "Customise your own style file"
 * checklist (embedded in the New Styles panel, toggled by AppHeader's header
 * button) - purely a presentation grouping, never consulted by
 * addDefaultStyles() itself. These three categories, and the order both they
 * and the styles within each of them render in, come directly from
 * STYLE-CATEGORIES.docx's own red section headings and item order (see the
 * extraction note on DEFAULT_STYLES below) - not chosen independently here. */
export type DefaultStyleCategory = 'Body text styles' | 'Heading Styles' | 'List styles'

export const DEFAULT_STYLE_CATEGORIES: DefaultStyleCategory[] = [
  'Body text styles',
  'Heading Styles',
  'List styles',
]

export interface DefaultStyleDefinition {
  name: string
  category: DefaultStyleCategory
  targetSignature: FormattingSignature
  kind: UserStyleKind
  listFormat: ListFormat
  listPreviewText?: string
}

/** StyleMash's bundled starter style set - extracted from
 * public/CLEAN-STYLES.docx, itself STYLE-CATEGORIES.docx (dropped at the
 * project root) with its informational-only "STYLE-TYPE"-styled section
 * labels stripped out. STYLE-CATEGORIES.docx lists each style as its own
 * paragraph, grouped under three red section headings - "Body text styles",
 * "Heading Styles", "List styles" - which is where `category` and
 * DEFAULT_STYLE_CATEGORIES' order come from, and both the category order and
 * the item order within each category are this array's order too, honoring
 * the source file exactly.
 *
 * Each signature is that document's own fully-cascaded look (through its
 * basedOn chain to its own docDefaults/theme) - the extraction used this
 * module's own resolveRunFormatting (the same per-run resolution the Style
 * Report itself uses, needed because one entry, "CRICOS/TEQSA", is direct
 * formatting layered on the Hyperlink character style rather than a named
 * style on its own) plus resolveStyleListFormat/buildStylePreviewMarker for
 * list items - the same functions referenceDocStyles.ts uses for Document B.
 * `category` is presentation-only (see DefaultStyleCategory) - it plays no
 * part in what addDefaultStyles() actually creates.
 *
 * `name` is each style's own real w:name (matching Word's own name, e.g.
 * "heading 1" lowercase) wherever a paragraph maps cleanly onto exactly one
 * named style; STYLE-CATEGORIES.docx deliberately departs from that in two
 * places - "HTML link" (renamed from the "Hyperlink" style's own technical
 * name for a clearer label) and "Heading 1 No Numbering" (a second, distinct
 * catalog entry deliberately sharing "Document title"'s own style/signature,
 * so its name has to come from the document's own item text instead) - both
 * of which are honored as-is rather than overridden back to a style's w:name.
 *
 * The whole document's docDefaults now sets Aptos (not Arial) as the base
 * font; the one style that still hardcoded a literal Arial override
 * (Hyperlink) had that override removed before extraction so nothing here
 * pins Arial - "HTML link"/"CRICOS/TEQSA" resolve with `fontFamily: null`
 * (no override), same as every heading. */
export const DEFAULT_STYLES: DefaultStyleDefinition[] = [
  {
    name: 'Normal',
    category: 'Body text styles',
    targetSignature: {
      fontFamily: 'Aptos',
      fontSizeHalfPt: 24,
      colorValue: 'auto',
      bold: false,
      italic: false,
      underline: null,
      strike: false,
    },
    kind: 'character',
    listFormat: 'none',
  },
  {
    name: 'Normal Bold',
    category: 'Body text styles',
    targetSignature: {
      fontFamily: 'Aptos',
      fontSizeHalfPt: 24,
      colorValue: 'auto',
      bold: true,
      italic: false,
      underline: null,
      strike: false,
    },
    kind: 'character',
    listFormat: 'none',
  },
  {
    name: 'caption',
    category: 'Body text styles',
    targetSignature: {
      fontFamily: 'Aptos',
      fontSizeHalfPt: 20,
      colorValue: 'auto',
      bold: false,
      italic: true,
      underline: null,
      strike: false,
    },
    kind: 'character',
    listFormat: 'none',
  },
  {
    name: 'HTML link',
    category: 'Body text styles',
    targetSignature: {
      fontFamily: null,
      fontSizeHalfPt: 24,
      colorValue: '467886',
      bold: true,
      italic: false,
      underline: 'single',
      strike: false,
    },
    kind: 'character',
    listFormat: 'none',
  },
  {
    name: 'CRICOS/TEQSA',
    category: 'Body text styles',
    targetSignature: {
      fontFamily: null,
      fontSizeHalfPt: 13,
      colorValue: '467886',
      bold: false,
      italic: false,
      underline: null,
      strike: false,
    },
    kind: 'character',
    listFormat: 'none',
  },
  {
    name: 'Document title',
    category: 'Heading Styles',
    targetSignature: {
      fontFamily: null,
      fontSizeHalfPt: 48,
      colorValue: '0F4761',
      bold: true,
      italic: false,
      underline: null,
      strike: false,
    },
    kind: 'character',
    listFormat: 'none',
  },
  {
    name: 'heading 1',
    category: 'Heading Styles',
    targetSignature: {
      fontFamily: null,
      fontSizeHalfPt: 40,
      colorValue: '0F4761',
      bold: true,
      italic: false,
      underline: null,
      strike: false,
    },
    kind: 'paragraph',
    listFormat: 'decimal',
    listPreviewText: '1',
  },
  {
    name: 'heading 2',
    category: 'Heading Styles',
    targetSignature: {
      fontFamily: null,
      fontSizeHalfPt: 32,
      colorValue: '0F4761',
      bold: true,
      italic: false,
      underline: null,
      strike: false,
    },
    kind: 'paragraph',
    listFormat: 'decimal',
    listPreviewText: '1.1',
  },
  {
    name: 'heading 3',
    category: 'Heading Styles',
    targetSignature: {
      fontFamily: null,
      fontSizeHalfPt: 28,
      colorValue: '0F4761',
      bold: true,
      italic: false,
      underline: null,
      strike: false,
    },
    kind: 'paragraph',
    listFormat: 'decimal',
    listPreviewText: '1.1.1',
  },
  {
    name: 'heading 4',
    category: 'Heading Styles',
    targetSignature: {
      fontFamily: null,
      fontSizeHalfPt: 24,
      colorValue: '0F4761',
      bold: false,
      italic: true,
      underline: null,
      strike: false,
    },
    kind: 'paragraph',
    listFormat: 'decimal',
    listPreviewText: '1.1.1.1',
  },
  {
    name: 'Heading 1 No Numbering',
    category: 'Heading Styles',
    targetSignature: {
      fontFamily: null,
      fontSizeHalfPt: 48,
      colorValue: '0F4761',
      bold: true,
      italic: false,
      underline: null,
      strike: false,
    },
    kind: 'character',
    listFormat: 'none',
  },
  {
    name: 'Heading 2 No Numbering',
    category: 'Heading Styles',
    targetSignature: {
      fontFamily: null,
      fontSizeHalfPt: 32,
      colorValue: '0F4761',
      bold: true,
      italic: false,
      underline: null,
      strike: false,
    },
    kind: 'character',
    listFormat: 'none',
  },
  {
    name: 'Heading 3 No Numbering',
    category: 'Heading Styles',
    targetSignature: {
      fontFamily: null,
      fontSizeHalfPt: 28,
      colorValue: '0F4761',
      bold: true,
      italic: false,
      underline: null,
      strike: false,
    },
    kind: 'character',
    listFormat: 'none',
  },
  {
    name: 'Heading 4 No Numbering',
    category: 'Heading Styles',
    targetSignature: {
      fontFamily: null,
      fontSizeHalfPt: 24,
      colorValue: '0F4761',
      bold: false,
      italic: true,
      underline: null,
      strike: false,
    },
    kind: 'character',
    listFormat: 'none',
  },
  {
    name: 'List Bullet',
    category: 'List styles',
    targetSignature: {
      fontFamily: 'Aptos',
      fontSizeHalfPt: 24,
      colorValue: 'auto',
      bold: false,
      italic: false,
      underline: null,
      strike: false,
    },
    kind: 'paragraph',
    listFormat: 'bullet',
    listPreviewText: '•',
  },
  {
    name: 'List Bullet 2',
    category: 'List styles',
    targetSignature: {
      fontFamily: 'Aptos',
      fontSizeHalfPt: 24,
      colorValue: 'auto',
      bold: false,
      italic: false,
      underline: null,
      strike: false,
    },
    kind: 'paragraph',
    listFormat: 'bullet',
    listPreviewText: '•',
  },
  {
    name: 'List Bullet 3',
    category: 'List styles',
    targetSignature: {
      fontFamily: 'Aptos',
      fontSizeHalfPt: 24,
      colorValue: 'auto',
      bold: false,
      italic: false,
      underline: null,
      strike: false,
    },
    kind: 'paragraph',
    listFormat: 'bullet',
    listPreviewText: '•',
  },
  {
    name: 'List Number',
    category: 'List styles',
    targetSignature: {
      fontFamily: 'Aptos',
      fontSizeHalfPt: 24,
      colorValue: 'auto',
      bold: false,
      italic: false,
      underline: null,
      strike: false,
    },
    kind: 'paragraph',
    listFormat: 'decimal',
    listPreviewText: '1.',
  },
  {
    name: 'List Number 2',
    category: 'List styles',
    targetSignature: {
      fontFamily: 'Aptos',
      fontSizeHalfPt: 24,
      colorValue: 'auto',
      bold: false,
      italic: false,
      underline: null,
      strike: false,
    },
    kind: 'paragraph',
    listFormat: 'decimal',
    listPreviewText: '1.',
  },
  {
    name: 'List Number 3',
    category: 'List styles',
    targetSignature: {
      fontFamily: 'Aptos',
      fontSizeHalfPt: 24,
      colorValue: 'auto',
      bold: false,
      italic: false,
      underline: null,
      strike: false,
    },
    kind: 'paragraph',
    listFormat: 'decimal',
    listPreviewText: '1.',
  },
]

/** Every DEFAULT_STYLES name, grouped by category and in DEFAULT_STYLES'
 * own order within each group - what the "Customise your own style file"
 * panel iterates to render its per-category dropdowns. */
export function groupDefaultStylesByCategory(): Map<DefaultStyleCategory, DefaultStyleDefinition[]> {
  const groups = new Map<DefaultStyleCategory, DefaultStyleDefinition[]>()
  for (const category of DEFAULT_STYLE_CATEGORIES) groups.set(category, [])
  for (const def of DEFAULT_STYLES) groups.get(def.category)!.push(def)
  return groups
}

/** Materializes DEFAULT_STYLES into `targetDocx` via the same "+ New Style"
 * code path (mergeStyles()/mergeParagraphStyle() with no source runs) every
 * other User-Created style goes through - so a default is byte-for-byte the
 * same shape, and behaves identically as a merge target, as any
 * manually-created or Document-B-imported style. A name collision with an
 * existing UserStyleRecord redefines that record's own styleId (newest
 * wins) instead of creating a duplicate - the same rule referenceDocStyles.ts
 * applies for Document B, so re-clicking "+ Defaults" after editing a
 * default's look just resets it rather than piling up a second entry.
 * Mutates targetDocx.stylesXml (and numberingXml, for a list default) in
 * place. Returns the full replacement for `existingUserStyles`.
 *
 * `enabledNames` restricts which of DEFAULT_STYLES actually get applied -
 * the "Customise your own style file" panel's checkboxes (see
 * useDocxWorkspace's enabledDefaultStyleNames) let a user opt individual
 * styles out before clicking "+ Defaults". Defaults to every style, so
 * existing callers/tests that don't care about customization are unaffected. */
export function addDefaultStyles(
  targetDocx: ParsedDocx,
  existingUserStyles: UserStyleRecord[],
  enabledNames: ReadonlySet<string> = new Set(DEFAULT_STYLES.map((d) => d.name)),
): UserStyleRecord[] {
  const existingByName = new Map(existingUserStyles.map((r) => [r.name, r]))
  const replacements = new Map<string, UserStyleRecord>()
  const brandNew: UserStyleRecord[] = []

  for (const def of DEFAULT_STYLES) {
    if (!enabledNames.has(def.name)) continue

    const collision = existingByName.get(def.name)
    const newStyleId =
      def.kind === 'character'
        ? mergeStyles(targetDocx, [], def.targetSignature, def.name, collision?.styleId)
        : mergeParagraphStyle(
            targetDocx,
            [],
            def.targetSignature,
            def.name,
            def.listFormat,
            collision?.styleId,
          )

    const record: UserStyleRecord = {
      styleId: newStyleId,
      name: def.name,
      targetSignature: def.targetSignature,
      kind: def.kind,
      listFormat: def.listFormat,
      listPreviewText: def.listPreviewText,
      createdAt: collision?.createdAt ?? Date.now(),
    }

    if (collision) {
      replacements.set(collision.styleId, record)
    } else {
      brandNew.push(record)
    }
  }

  const merged = existingUserStyles.map((r) => replacements.get(r.styleId) ?? r)
  return [...merged, ...brandNew]
}
