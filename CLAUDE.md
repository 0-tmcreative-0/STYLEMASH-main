# CLAUDE.md

Working notes for Claude Code in this repo. **Every change made to this project
gets an entry in the [Change log](#change-log) at the bottom of this file.**

## What this is

StyleMash — a local-only, client-side browser tool that standardizes text
formatting in Word documents. A `.docx`/`.dotx` is parsed in the browser
(JSZip + `DOMParser`), every distinct *resolved* text appearance is listed,
and the user folds groups of them into clean named styles, then downloads the
result. Nothing is uploaded anywhere. See `README.md` for the user-facing docs.

## Commands

| Command | Notes |
|---|---|
| `npm run dev` | Vite dev server — **the documented primary way users run this app**, so dev-only behaviour (React StrictMode!) is production behaviour here. |
| `npm run build` | `tsc -b` (app + node + **tests**) then `vite build`. |
| `npm test` | Vitest, jsdom environment. |
| `npm run lint` | Oxlint. |

## Architecture

- `src/lib/ooxml/` — all OOXML logic, framework-free and directly unit-tested.
  This is where the real complexity lives.
- `src/hooks/useDocxWorkspace.ts` — the single reducer owning all workspace
  state (parsed doc, style report, user styles, selection, modals, undo).
- `src/components/` — presentational; no OOXML knowledge beyond reading a
  `FormattingSignature`.
- `tests/` — mirrors `src/lib/ooxml/` plus two React tests.

### Load-bearing invariants

Break any of these and the output document is quietly corrupt — Word's "needs
repair" prompt is the usual symptom.

1. **The reducer must stay pure.** `documentXml`/`stylesXml`/`numberingXml` are
   mutated *in place* all session long (so live `RunRef` Element pointers stay
   valid), but that mutation happens in `useDocxWorkspace`'s **action
   creators**, never in the reducer. React invokes a reducer more than once per
   dispatch — StrictMode does it on *every* dispatch — so a mutating reducer
   applies each merge twice. Action creators read current state through
   `stateRef`, do the work, and dispatch the finished result.
2. **`<w:rPr>` children have a fixed schema order** (`RPR_CHILD_ORDER` in
   `constants.ts`). Always insert via `insertRPrChildInOrder()`; never
   `appendChild` a new rPr child.
3. **`w:rStyle` may only point at a character style; `w:pStyle` only at a
   paragraph style.** Anything dispatching to a merge must branch on
   `UserStyleRecord.kind` (`mergeStyles` vs `mergeParagraphStyle`).
4. **A `<w:p>` can contain other `<w:p>`s** (text boxes:
   `w:drawing → … → w:txbxContent → w:p`). Attribute a run to its *nearest*
   ancestor paragraph — use `getOwnRuns()` from `styleReport.ts`, not a raw
   `getElementsByTagNameNS(NS.w, 'r')` descendant scan.
5. **A brand-new `word/numbering.xml`** (created for a document's first list
   style) must also be registered in `[Content_Types].xml` *and*
   `word/document.xml.rels`, or Word ignores the part entirely. See
   `serializeDocx.ts#ensureNumberingPartRegistered`.
6. **`JSZip#clone()` shares its `.files` table by reference.** Use
   `contentMerge.ts#cloneZipForOutput` when producing a zip derived from
   another live `ParsedDocx`.

### Deliberate v1 scope limits (documented, not bugs)

Headers/footers/footnotes aren't scanned; `w:themeTint`/`w:themeShade` aren't
applied; toggle properties use last-writer-wins rather than the spec's XOR
semantics; paragraph-mark run properties (`w:pPr/w:rPr`) are ignored; the
content-merge path doesn't reconcile `w:numId` or relationship ids.

### Features currently switched off in the UI

Both are fully wired up (state, handlers, dialogs, tests) but hidden behind a
hardcoded `false`, awaiting more work:

- **Bulk-match to Document B** — `src/components/StyleReportPanel.tsx`, the
  `{false && hasReferenceStyles && …}` block. This is also the one thing
  Oxlint warns about (`no-constant-binary-expression`).
- **Merge content into Document B** — `CONTENT_MERGE_ENABLED` in
  `src/components/DocumentPreviewPanel.tsx`.

## Gotchas for a fresh session

- **This became a git repository mid-session on 2026-09-18** (single
  "Initial commit"). Earlier advice in this file assumed no history existed —
  that's stale; `git log`/`git diff` now work normally. Commit/push only when
  the user asks, per standing instructions.
- **`npm run lint` fails on an odd-numbered Node** (19/21/23) with
  `Cannot find module '@oxlint/binding-darwin-arm64'`. Oxlint's native binding
  is an optional dependency npm skips on an unsupported engine. Fix: install
  an LTS release (18/20/22) and reinstall `node_modules` under it — e.g. on
  this machine, `brew install node@20 && brew unlink node && brew link
  --overwrite node@20`, then `rm -rf node_modules && npm install`. A
  same-major-version symlink under `/opt/homebrew/opt/node@20` is not proof
  Node 20 is actually installed — check `readlink` before trusting it; it can
  point straight back at the odd-numbered Cellar keg if only the default
  `node` formula was ever installed.
- **"Edit XML" (`XmlEditorModal`, `useDocxWorkspace.ts#openXmlEditor`) has no
  UI entry point.** The reducer, action, undo support, and modal are all
  intact and covered by tests, but no component ever calls
  `actions.openXmlEditor()` — grep confirms it. Don't describe this as
  available to end users until something wires it up.

---

## Change log

### 2026-09-18 — Code review pass, then fixes

Full read-through of the codebase (no prompt history was saved from the
original build). Three correctness bugs found, each reproduced with a failing
test before being fixed. All 70 tests pass; `npm run build` and `npm run lint`
are clean.

#### Fixed: every merge silently created a duplicate, orphaned `<w:style>`

`src/hooks/useDocxWorkspace.ts` — the reducer called `mergeStyles()` /
`mergeParagraphStyle()` / `addDefaultStyles()` / `removeStyleById()` /
`applyXmlFragmentToRunRefs()` directly, i.e. it performed side effects. React
StrictMode (enabled in `src/main.tsx`) invokes a reducer twice per dispatch, so
every merge ran twice: the second run saw the style the first had just created,
generated a colliding-but-distinct id (`MyStyle` → `MyStyle1`), and left the
first definition behind unreferenced. A paragraph/list merge also created a
second orphaned `<w:abstractNum>`/`<w:num>` pair each time. *"+ Defaults"* added
34 style definitions instead of 17. Because `npm run dev` is the documented way
to run this app, every saved document was affected.

Restructured so the reducer is pure: all mutating work moved into the action
creators (which run exactly once per gesture and read current state via a
`stateRef`), with the reducer only storing already-computed results. Actions
renamed to past-tense result actions (`MERGE_APPLIED`, `XML_EDIT_APPLIED`,
`BULK_MERGE_APPLIED`, `DEFAULT_STYLES_ADDED`, `USER_STYLES_CLEARED`,
`UNDO_APPLIED`, `REFERENCE_DOC_REMOVED`, plus a shared `ACTION_FAILED`).

This also removed the `pristineSnapshotCache` `WeakMap`, which existed purely to
work around the same double-invocation for undo snapshots — with the mutations
out of the reducer it has nothing left to guard.

#### Fixed: bulk-merge applied a paragraph style through `w:rStyle`

`src/lib/ooxml/bulkMergeMatchedStyles.ts` — always called `mergeStyles()`
(the character-style path), even for a `kind: 'paragraph'` record. A Document B
list style materializes as exactly that (`referenceDocStyles.ts`), so merging
into one wrote `<w:rStyle w:val="…">` pointing at a
`<w:style w:type="paragraph">` — invalid OOXML that Word can't resolve — while
leaving the paragraph on its old `w:pStyle` and old list. Now branches on
`record.kind`, matching what the single-target "Merge N selected here" flow
already did. (Reachable only once the bulk-match UI is switched back on.)

#### Fixed: text-box runs counted and rendered twice

`src/lib/ooxml/styleReport.ts`, `src/components/DocumentPreviewPanel.tsx` — both
walked `paragraphEl.getElementsByTagNameNS(NS.w, 'r')`, a descendant scan. A
`<w:p>` containing a text box nests whole paragraphs inside itself, so every run
in a text box was attributed to both its own paragraph and the anchor
paragraph: inflated occurrence counts, double-rendered preview text, and the
same run merged twice. Added `getOwnRuns()`, which keeps the descendant scan
(so `w:hyperlink`/`w:ins`/`w:sdt`-wrapped runs are still found) but filters to
runs whose nearest ancestor `<w:p>` is the paragraph being processed.

#### Fixed: "Edit XML" was not undoable

`applyXmlEdit` mutated the document without pushing an undo snapshot, so the
Undo button skipped straight past it to the previous merge. It now snapshots
like every other mutating action.

#### Type-checking gaps closed

- `tsconfig.app.json` had no `"strict"` — `strictNullChecks` et al. were off for
  the whole app. `src/` turned out to be strict-clean already, so `"strict":
  true` is now on at zero cost.
- `tests/` was outside every tsconfig project, so `npm run build` never
  type-checked it. Added `tsconfig.test.json` and referenced it from the root
  project. This immediately caught five `UserStyleRecord` fixtures missing the
  required `kind`/`listFormat` fields (in `bulkMergeMatchedStyles.test.ts`,
  `referenceDocStyles.test.ts`, `styleReport.test.ts`); all corrected.

#### Tests added (60 → 70)

- `tests/workspaceReducer.test.tsx` (new, 8 tests) — mounts `useDocxWorkspace`
  inside `<StrictMode>`, exactly as `main.tsx` does, and asserts a character
  merge, a list merge and `"+ Defaults"` each produce exactly one definition;
  plus undo of a merge, undo of an XML edit, and inline merge-failure handling.
- `tests/styleReport.test.ts` — a run inside a `w:txbxContent` text box is
  counted once.
- `tests/bulkMergeMatchedStyles.test.ts` — a paragraph-kind reference record is
  applied via `w:pStyle`, never `w:rStyle`.

#### Not changed (flagged for a decision)

- The two features hidden behind a hardcoded `false` (see above) — left as-is;
  deleting or re-enabling them is a product call.
- README/UI terminology drift (see above).
- `npm run lint` failing under Node 21 — an environment issue, not a repo one.
  Lint *was* verified during this pass by temporarily installing the missing
  `@oxlint/binding-darwin-arm64`: it reports nothing beyond the known
  `no-constant-binary-expression` warning on the intentionally-disabled
  bulk-match block. That install also pulled newer minor versions of several
  deps into `node_modules` (it bypassed the lockfile), so `npm ci` was run
  afterwards to restore the exact locked versions — `node_modules` now matches
  `package-lock.json` again and lint is back to its pre-existing broken state
  on Node 21. `package.json` and `package-lock.json` were never modified.

### 2026-09-18 — Follow-up: audited the two disabled features, fixed the environment, synced the README

#### Audited the two hidden features (kept disabled, as instructed)

Reviewed `contentMerge.ts`/`ContentMergeDialog.tsx` and the bulk-match block in
`StyleReportPanel.tsx` for correctness issues beyond the one already fixed.
Nothing else rose to "obvious bug"; one real UX rough edge found and left
alone since the feature stays off:

- The bulk-match checkbox (`{false && hasReferenceStyles && …}` block,
  `StyleReportPanel.tsx`) only *adds* matching variants to the selection when
  checked (`onSelectMatchingReferenceStyles`) — unchecking it does nothing, so
  the checkbox's visual state stops reflecting what's actually selected. Minor,
  and moot while the block is unreachable; worth a second look if it's ever
  switched back on.

Separately (not one of the two flagged features, but found while checking
UI wiring for the README): **"Edit XML" has no UI entry point at all** — see
the new Gotchas entry above. Unlike the other two, this isn't gated by a
`false` flag; the trigger was simply never wired to any component. Left as-is
per the same "don't touch dead code speculatively" reasoning, but the README
can no longer document it as a working feature (see below).

#### Fixed: Node 21 → Node 20, for real this time

The earlier "verified lint passes on Node 20" claim in the previous entry was
wrong. What looked like a Node 20 binary — `/opt/homebrew/opt/node@20/bin/node`
— was a stale symlink to `../Cellar/node/21.7.3`; `node@20` was never actually
installed, so that "test" silently ran Node 21 again. Actually installed it
(`brew install node@20`), linked it as the active `node`
(`brew unlink node && brew link --overwrite node@20` — `node` now resolves to
v20.20.2), and reinstalled `node_modules` under it (`rm -rf node_modules &&
npm install`, which respected the lockfile: React/JSZip versions unchanged).
`npm run lint` is now genuinely clean — nothing but the expected
`no-constant-binary-expression` warning on the disabled bulk-match block.
`npm test` (70/70) and `npm run build` still pass.

#### README brought in line with the actual UI

`README.md`'s "Using the app" section was written against an earlier UI: it
said *"Style Report"* / *"User-Created Styles"* / *"Do it"* / *"Save
locally"*, none of which match current copy (*"Current styles"*, *"New
Styles"*, *"Mash it"*, *"Save your file"*), and it documented **"Edit XML"**
as a working step even though (see above) there's no button anywhere that
opens it. Rewrote the numbered walkthrough to match current UI copy exactly,
dropped the Edit XML step, and added coverage for four real, working features
the old README never mentioned at all: **+ Defaults**, merging straight into
a target style via **New Styles**' "Merge N selected here", **Attach Document
B**, and **Undo**. The rest of the README (intro paragraph, requirements,
scripts table, troubleshooting) was checked against current behavior and
left untouched — it was already accurate.

Not addressed (out of scope for this pass, flagged for later):

- `HelpModal.tsx`'s placeholder copy has the same terminology drift the README
  had (*"New styles"* button, *"Do it"* button) — same root cause, different
  file. Not touched since only the README was in scope.
- `AppFooter.tsx` shows `v.0.0.2-alpha`; `package.json` says `"version":
  "0.0.2"` (no `-alpha`). Cosmetic version-string mismatch, not touched.

### 2026-09-18 — Fixed 2 moderate `npm audit` vulnerabilities

`npm install` reported 2 moderate-severity advisories, both the same root
cause: `@vitest/mocker` (vitest's mocking layer, pulled in transitively by the
`vitest` devDependency) versions 2.1.0–4.1.10 have a path-traversal /
arbitrary-file-read bug via redirected mocks
([GHSA-82fw-gwwq-j7x9](https://github.com/advisories/GHSA-82fw-gwwq-j7x9)).
Dev-only exposure (vitest never ships in `dist/`), but real: anything that can
influence a mock path while `npm test` runs could read arbitrary files.

There is no patched 3.x release — 3.2.7 was the last one; the fix landed in
4.1.11. Bumped `package.json`'s `vitest` from `^3.2.7` to `^4.1.11` (the
minimal patched version, not the newer `5.x` line — the repo's `vite@6.4.3`
satisfies vitest 4's peer range of `^6.0.0 || ^7.0.0 || ^8.0.0`, and
`@types/node@24.13.3` satisfies its `@types/node` peer range too, so no other
version bumps were needed). Ran `npm install` (no `--force`) to pick up the
new resolution — `package-lock.json` updated accordingly.

Verified: `npm audit` → 0 vulnerabilities; `npm test` → 70/70 still pass
(vitest's `RUN v4.1.11` banner is the only visible difference — no config or
test-file changes were needed); `npm run build` and `npm run lint` unaffected.

### 2026-09-18 — Real bundled default styles from a reference file; customize-defaults panel; layout moves

Version bumped `0.0.2` → `0.1.0` (`package.json` + the footer's `v.0.1.0-beta`,
was `v.0.0.2-alpha`) - the scope here is a real feature release, not a patch.

#### DEFAULT_STYLES now comes from a real reference file, not hand-typed data

The user dropped `REF-DOC/CLEAN-STYLES.docx` at the project root - a from-scratch
recreation of the original deleted TEST-DOC reference (see the 2026-08 history
this file no longer has, but `defaultStyles.ts`'s old doc comment referenced).
Extracted its actually-applied styles the same way `referenceDocStyles.ts`
extracts Document B's, via a disposable Vitest script (JSZip + jsdom's
DOMParser, no browser needed) driving this codebase's own
`resolveStyleRPr`/`trackedChildrenToSignature`/`resolveStyleListFormat`/
`buildStylePreviewMarker` - never hand-transcribed by reading raw OOXML. Every
resolved signature came out byte-identical to the old hardcoded set except one
rename (the old set's "Heading 1 No Numbering" is now "Document title" -
same signature, same role) - strong evidence CLEAN-STYLES.docx is a faithful
successor to the file that produced the original data.

One gap the "styles referenced by an explicit w:pStyle/w:rStyle" extraction
heuristic missed entirely: **"Normal"** itself, since every unstyled paragraph
in the reference file uses it *implicitly* (no w:pStyle at all) - the user
caught this in testing ("it is important that users can merge source files to
the 'Normal' style"). Added by hand via the same resolution functions, scoped
to styleId "Normal" directly. `DEFAULT_STYLES` is now 18 entries, not 17 -
guarded by a regression test (`tests/defaultStyles.test.ts`) so a future
re-extraction pass can't silently drop it again the same way.

`public/CLEAN-STYLES.docx` now ships in every build (Vite serves `public/` at
the root, confirmed present in `dist/` after `npm run build`) - both as the
extraction's data source and as a real download for users (see below). The
`REF-DOC` folder was removed once both were in place, per the user's own
instruction.

#### "Customise your own style file" - new header button + modal

New button in `AppHeader.tsx`, next to Help, **gated on a document being
loaded** (reuses the existing `filename` prop as the signal - it edits what
"+ Defaults" applies, and that button doesn't exist on the upload screen; the
user caught this too, in testing). Opens `CustomizeDefaultStylesModal.tsx`:

- `DEFAULT_STYLES` grouped by a new presentation-only `category` field
  (`DefaultStyleCategory`: Headings / Bulleted Lists / Numbered Lists / Text
  Styles - see `defaultStyles.ts#groupDefaultStylesByCategory`) into
  independently-collapsible sections, each with a live count ("N of M
  selected") and a chevron that rotates on expand/collapse. The
  expand/collapse animation is a CSS grid-template-rows trick
  (`grid-rows-[0fr]` → `grid-rows-[1fr]`) rather than a JS-measured
  max-height - animates to the content's real height with no measurement.
- A checkbox per style (see FaCheckbox below) toggles whether "+ Defaults"
  applies it - state lives in `useDocxWorkspace` as `enabledDefaultStyleNames`,
  filtered into `addDefaultStyles()`'s new third parameter
  (`enabledNames: ReadonlySet<string>`, defaulting to "every style" so
  existing callers/tests are unaffected).
- **Deliberately its own `useState`, not reducer state**: every mutating
  reducer action spreads `initialState` wholesale on `FILE_LOADED`/`RESET`
  (see the reducer's existing pattern), which would silently re-check
  everything on "Mash a different file" if this lived there. It's a user
  preference, not document state - confirmed with a regression test
  (`enabledDefaultStyleNames survives "Mash a different file" (reset)`).
- A "Download reference style file" link to `/CLEAN-STYLES.docx` - the other
  half of "make this file available as a download for users". A plain
  `<a href download>` to a `public/` asset, not an Artifact - the Artifact
  tool's "no download links" sandbox restriction doesn't apply to this real
  web app.

#### FaCheckbox - a shared checkbox glyph, mirrored into the Current Styles list

New `src/components/FaCheckbox.tsx`: a real (visually hidden, `sr-only`)
`<input type="checkbox">` for keyboard/screen-reader support, with a
FontAwesome glyph layered on top for the visible state - regular-style
(outline) square for unchecked, solid check-square for checked. Worth noting:
`@fortawesome/free-solid-svg-icons`' `faSquare` is a *filled* square, not an
outline - using it for "unchecked" renders as an ambiguous solid gray blob
with no visible border, not an empty box (caught by screenshot-testing in
Chrome, not by any unit test - there's no automated check for "does this
icon look like a checkbox"). Pulled in
`@fortawesome/free-regular-svg-icons` for the outline variant instead.

Used by both `CustomizeDefaultStylesModal` and - per the user's explicit
request to mirror the customize panel's checkbox style back into the Current
Styles list - `StyleVariantRow.tsx`, replacing its previous fully-invisible
`sr-only`-only checkbox (selection there was previously indicated by row
background color alone). The Current Styles list itself stays a flat,
non-grouped list as instructed ("don't put these style into drop downs, the
expanded view is helpful") - only the checkbox glyph is shared, not the
dropdown grouping.

New dependencies: `@fortawesome/fontawesome-svg-core`,
`@fortawesome/free-solid-svg-icons`, `@fortawesome/free-regular-svg-icons`,
`@fortawesome/react-fontawesome` (all pinned exact versions, all bundled by
Vite - no CDN fetch, consistent with this being a local-only, offline-capable
tool). `npm audit`: 0 vulnerabilities.

#### Layout: Undo and Save moved out of Document Preview's header

Per explicit instruction: **Undo** moved from `DocumentPreviewPanel`'s header
to `StyleReportPanel`'s footer, immediately left of "Mash it" (same row,
`UndoButton` sized to content next to a `flex-1` "Mash it"). **Save your
file** moved to its own full-width row below that, in the same footer
(`SaveButton` gained a `className` prop for this). `DocumentPreviewPanel`'s
header now holds only the title + tooltip, as instructed ("leave the
document Preview window header otherwise as-is"). Both button components'
own doc comments (which named their old locations) were updated to match.

Follow-up the user caught in testing: removing those buttons left
`DocumentPreviewPanel`'s header shorter than `StyleReportPanel`/
`UserStylesPanel`'s (their headers still had button rows forcing more
height). Added `min-h-15` (60px) to all three panel headers so they align
regardless of what each one's header row happens to contain - measured via
Chrome DevTools to confirm 59px→60px→60px before/after, not eyeballed.

#### Small copy changes, also from live testing feedback

- "Attach Document B (optional)" → **"Attach style file (recommended)"**
  (`AttachReferenceDocButton.tsx`; a stale doc comment in `UserStylesPanel.tsx`
  naming the old label was also updated).
- Footer version string `v.0.0.2-alpha` → `v.0.1.0-beta` (see version bump
  above).

#### Tests added (70 → 84)

- `tests/defaultStyles.test.ts` (new) - every `DEFAULT_STYLES` name is
  unique and categorized; "Normal" is present with the right kind/listFormat
  (the regression guard mentioned above); `groupDefaultStylesByCategory`
  partitions the full set correctly; `addDefaultStyles`' new `enabledNames`
  parameter filters correctly (including the empty-set and
  omitted-parameter/back-compat cases) and is purely additive across repeated
  calls (never retroactively removes a style already created).
- `tests/faCheckbox.interaction.test.tsx` (new) - the real checkbox input
  exists and is `sr-only`; clicking it fires `onToggle` exactly once; the
  correct icon (`square-check` vs `square`) renders for each state.
- `tests/workspaceReducer.test.tsx` - `enabledDefaultStyleNames` starts with
  every default checked; unchecking one excludes it from "+ Defaults"; the
  enabled-set survives `reset()` (the "preference, not document state" guard
  described above).

#### Verified end-to-end in Chrome (dev server, not just unit tests)

Uploaded `CLEAN-STYLES.docx` itself as a live test document and walked the
whole new flow: header button appears only once a document is loaded and
disappears again on "Mash a different file"; category sections
expand/collapse with the animation; unchecking "caption" and clicking Done
made `+ Defaults` produce 17 records instead of 18, confirming the filter is
wired end-to-end (not just at the reducer level); selecting "Normal" as a
merge target and clicking "Merge 1 selected here" then **Undo** correctly
applied and then fully reverted the merge (style removed, occurrence count
restored, progress bar restored) in the button's new footer location; no
JS console errors beyond an unrelated Chrome-extension messaging warning.

### 2026-09-18 — Fixed: the grouped checklist was never actually in the New Styles panel

Immediate user correction on the previous entry's work: "Customise your own
style file" opened a floating centered modal (`CustomizeDefaultStylesModal.tsx`)
rather than living inside the New Styles panel itself, which is what the
original request actually asked for (the request's own later line - "ensure
any checkbox additions employed in the New Styles list are mirrored in the
Current Styles list" - only makes sense if the checklist genuinely is part of
the New Styles list, not a separate popup).

Restructured: deleted the modal, extracted its content into
`DefaultStylesChecklist.tsx` (no dialog chrome - just the download link +
category sections), and embedded it directly inside `UserStylesPanel.tsx`
between the header and the styles list, collapsed/expanded via the same
grid-template-rows animation its own category sections use. AppHeader's
button no longer owns a modal's open state - it's now a plain toggle
(`isCustomizeOpen`, lifted to `App.tsx` as local UI state, a sibling of both
components) with an active/pressed visual state (`aria-pressed` + an indigo
fill while open) so it reads as "this button controls that panel section",
not "click for a popup".

Caught two rendering issues only visible in the browser, not in any unit
test: (1) a stale Vite HMR module instance made the header button's `isOpen`
visual state disagree with the actual panel section's rendered state after a
few incremental edits - resolved by a hard navigation, not a code fix, and a
reminder that HMR staleness is a real failure mode worth ruling out before
debugging app logic; (2) none, once past that - `npm run build`/`test`/`lint`
all stayed green throughout since none of this touched the parts those cover
(state plumbing and DEFAULT_STYLES data), only which component renders it.

#### Categories re-cut: from list-type splits to what merging actually needs

Also per direct instruction: `DefaultStyleCategory` changed from `'Headings' |
'Bulleted Lists' | 'Numbered Lists' | 'Text Styles'` to `'Headings' | 'Normal'
| 'Lists' | 'Text Styles'` - every list (bulleted or numbered) now groups
together instead of splitting by type, and "Normal"/"Normal Bold" get their
own dedicated group instead of sitting in the catch-all "Text Styles" bucket
with caption/Hyperlink - reflecting that "Normal" is the single most
important merge target for real documents (see the previous entry's own note
on why it was added at all), not just one text style among several. No test
had to change: `tests/defaultStyles.test.ts`'s assertions are all generic
over `DEFAULT_STYLE_CATEGORIES`/`groupDefaultStylesByCategory`, never
hardcoding a category name.

#### Checklist rows now genuinely match Current Styles' presentation

Per direct instruction ("match the Current Styles presentation"), rebuilt
each checklist row (new `DefaultStyleRow` in `DefaultStylesChecklist.tsx`) to
the same shape `StyleVariantRow` uses in the Current Styles panel - FaCheckbox
+ a live-styled sample line (`signatureToCss`, with the list marker prefixed
when present) + a `describeSignature()` line underneath - rather than the
single compact line it had before. The two things a bundled default
genuinely has no equivalent of were left out rather than faked: an origin
line (nothing to attribute a definition to) and an occurrence badge (nothing
counted yet).

Re-verified end-to-end in Chrome after the restructure: the button's
active/inactive state, the panel section actually expanding in place (not a
modal), all four categories (Headings 8/8, Normal 2/2, Lists 6/6, Text Styles
2/2) present with correct counts, and the "Normal" category's rows rendering
in the new StyleVariantRow-matching shape. `npm run build`/`test` (84/84)/
`lint` all still pass.

### 2026-09-18 — DEFAULT_STYLES re-extracted from a new reference file with real
category grouping; Normal's default font moves to Aptos; "BETA" tag added

#### DEFAULT_STYLES/categories now come from STYLE-CATEGORIES.docx, not a
hand-assigned `category` field

The user dropped `STYLE-CATEGORIES.docx` at the project root: the same idea
as the original CLEAN-STYLES.docx (one paragraph per bundled style), but this
time genuinely organized under three red section headings - "Body text
styles", "Heading Styles", "List styles" - which is what the "Customise your
own style file" checklist should have been grouping by all along instead of
the previous pass's own invented categories (Headings/Normal/Lists/Text
Styles). A fourth style, "STYLE-TYPE" (basedOn "Document title", colored red
`EE0000`), was applied to the three section-heading paragraphs themselves and
explicitly called out as informational-only scaffolding to delete once done.

Extraction this time couldn't just resolve each item's *named style* in
isolation the way the original CLEAN-STYLES.docx pass did (see that entry
above) - one entry, **"CRICOS/TEQSA"**, turned out to be direct formatting
(`b="0"`, `sz="13"`, `u="none"`) layered on top of the "Hyperlink" character
style at the paragraph itself, with no named style of its own, so resolving
just "Hyperlink" in isolation would have produced the wrong (full-size, bold,
underlined) signature. Switched the extraction to `resolveRunFormatting` -
the exact per-run cascade (docDefaults → paragraph style → character style →
direct overrides) the Style Report itself uses - run once per item's
representative paragraph, which handles both the 18 clean named-style cases
and this one correctly with no special-casing. `resolveStyleListFormat`/
`buildStylePreviewMarker` (per the referenced pStyle) still supply
listFormat/listPreviewText for list items, same as before.

Cleanup before extraction (via a disposable Vitest script, deleted once its
job was done, same as the original CLEAN-STYLES.docx pass): removed the
leading blank paragraph, the three STYLE-TYPE-styled/red-colored section
label paragraphs, and the "STYLE-TYPE" style definition itself from
styles.xml. The cleaned result became the new `public/CLEAN-STYLES.docx`
(both the extraction source and the "Download reference style file" asset,
same dual role as before), and `STYLE-CATEGORIES.docx` was deleted from the
project root per instruction once both were in place.

`DEFAULT_STYLES` is now 20 entries (was 18): two are genuinely new -
**"CRICOS/TEQSA"** (above) and **"Heading 1 No Numbering"**, a second catalog
entry deliberately sharing "Document title"'s own style/signature (distinct
conceptual role, identical look - the same relationship the *previous*
CLEAN-STYLES.docx pass's own doc comment described in reverse, where
"Heading 1 No Numbering" had been *folded into* "Document title"; this
source file restores it as its own entry) - and one is a deliberate rename:
the "Hyperlink" style's item is labeled **"HTML link"** in this source file
rather than its own technical style name, honored as-is since it's clearly
intentional (unlike, say, "List bullet" vs. "List Bullet 2/3"'s inconsistent
capitalization elsewhere in the same file, which reads as an unintentional
typo - `name` for every other item still comes from that item's own real
`w:name`, e.g. lowercase "heading 1"/"caption", matching the previous pass's
own convention).

`DefaultStyleCategory` is now `'Body text styles' | 'Heading Styles' | 'List
styles'` (was `'Headings' | 'Normal' | 'Lists' | 'Text Styles'`), and
`DEFAULT_STYLE_CATEGORIES`'s order - which is also the checklist's own
rendering order - matches the source file's own section order exactly (Body
text styles → Heading Styles → List styles), per explicit instruction to
honor the source file's order for both categories and the items within each
one. `DefaultStylesChecklist.tsx` needed no code change - it was already
fully generic over `DefaultStyleCategory`/`DEFAULT_STYLE_CATEGORIES` - only
its own doc comment (which named the old four categories) was updated.
`tests/defaultStyles.test.ts` also needed no change: its assertions were
already generic over the category set rather than hardcoding one.

#### Normal's default font moves from Arial to Aptos

Per explicit instruction, checked for any remaining "Arial" reference in the
data or downloadable file this pass produces. STYLE-CATEGORIES.docx's own
`docDefaults`/"Normal" style now sets Aptos (not Arial) - correct and left
alone - but its "Hyperlink" character style still had a literal, stale
`<w:rFonts w:ascii="Arial" w:hAnsi="Arial"/>` override left over from an
earlier era of the file. Removed that override (in the same cleanup pass
above) before extraction, so "HTML link"/"CRICOS/TEQSA" now resolve with
`fontFamily: null` (no override - inherits whatever body font the style ends
up merged into, the same as every heading already did) rather than a
hardcoded font. Verified via the extraction's own JSON dump and by grepping
the cleaned `public/CLEAN-STYLES.docx`'s styles.xml: zero `fontFamily:
'Arial'` entries anywhere in the new `DEFAULT_STYLES`, and no literal Arial
left on any of the 20 extracted items' style chains. (Several unrelated,
never-applied auto-generated "*Char" linked styles elsewhere in styles.xml
still say Arial - out of scope, since nothing in the app's own 20 bundled
items or their cascades ever reads them.)

#### "BETA" tag added next to the app name

Small addition to `AppHeader.tsx`: a green "BETA" label next to "StyleMash"
in the header - purely cosmetic, no state or behavior involved. First pass
used `font-variant: small-caps` on "Beta"; per follow-up feedback, switched
to genuinely-uppercase text (`uppercase`, simpler and equally legible) at a
smaller size (`text-[10px]`, vs. "StyleMash"'s `text-xl`), and moved both
into their own nested `flex items-baseline` span (separate from the outer
`items-center` row the mark/wordmark/tag all still sit in) so "BETA"'s
baseline lines up with "StyleMash"'s rather than being vertically centered
against it.

### 2026-09-18 — Style preview lists now render with the same font fallbacks as
the Document Preview

Per direct feedback: every sample-line preview outside `DocumentPreviewPanel`
(Style Report rows via `StyleVariantRow.tsx` and its own multi-variant group
header in `StyleReportPanel.tsx`, New Styles rows in `UserStylesPanel.tsx`,
the Customise checklist's rows in `DefaultStylesChecklist.tsx`, and the Merge
Dialog's live preview) called `signatureToCss()` directly, which renders a
signature's `fontFamily` verbatim with no fallback. `DocumentPreviewPanel`
alone wrapped that same call in its own local `previewCss()`, layering
standard web-safe fallbacks (`Calibri, "Segoe UI", Arial, sans-serif`) after
the real font name - added originally because most custom/corporate fonts
aren't installed in a browser. That meant every list *other than* the
Document Preview would silently fall back to the browser's own arbitrary
default font instead, so the same style could look different from panel to
panel - most visibly now that the bundled default font is Aptos (see the
prior entry), which is essentially never installed outside Microsoft 365.

Moved the fallback-chain logic into `signatureToCss()` itself
(`src/lib/signatureToCss.ts`) so every call site gets it for free, and
deleted `DocumentPreviewPanel`'s now-redundant local `previewCss()` wrapper
(along with its since-unused `FormattingSignature` import) in favor of
calling `signatureToCss()` directly, same as everywhere else. No call site
needed a change beyond that deletion - `signatureToCss`'s signature and
return shape (`CSSProperties`) are unchanged, only its `fontFamily` value's
contents.

Verified in Chrome: uploaded `public/CLEAN-STYLES.docx` and zoomed into the
"Normal" row in Current Styles alongside the "Normal" line in Document
Preview - both now render in the same Calibri fallback glyph shapes (Aptos
isn't installed on this machine), where before the Style Report list row
would have rendered in the browser's own default font instead. Also
spot-checked the Customise checklist's rows post-fallback. `npm run
build`/`test` (84/84, no test referenced `signatureToCss`'s/`previewCss`'s
output directly so none needed changes) / `lint` all still pass.

### 2026-09-18 — New Styles rows get the same checkbox as Current Styles, with
a deliberately different highlight color

Per direct request: added `FaCheckbox` to each row in `UserStylesPanel.tsx`
("New Styles"), the same glyph `StyleVariantRow.tsx` ("Current Styles") uses
- `checked={isTarget}`, `onToggle` calling the same
`onToggleSelectTarget(record.styleId)` the row's own `onClick` already did,
so clicking the checkbox and clicking anywhere else on the row do the exact
same thing (as they already did for Current Styles' rows - `FaCheckbox`'s
inner input stops click propagation and fires `onChange` once, so there's no
double-toggle from the click bubbling to the `<li>` too).

Note this doesn't make New Styles genuinely multi-select the way Current
Styles is (`selectedTargetStyleId` is still a single nullable string, not a
`Set`, since only one style can ever be a merge *target*) - the checkbox here
is purely a visual affordance matching the request, layered on top of the
existing single-select "click a row to pick a merge target" behavior, not a
change to that behavior itself.

Also per direct request, the selected-row highlight now uses amber
(`bg-amber-200`/`border-l-amber-500`, was `bg-indigo-200`/
`border-l-indigo-500` - identical to Current Styles' own selected-row color)
rather than reusing Current Styles' indigo, so a glance at both panels shows
two different "picked" states rather than what could read as one shared
selection between unrelated lists. Deliberately loud/temporary per
instruction ("something visually noticeable for now... fix later") - flagged
here as a placeholder, not a final color decision.

Verified in Chrome: uploaded `public/CLEAN-STYLES.docx`, ran "+ Defaults",
confirmed every New Styles row now shows an unchecked checkbox, clicking a
row's checkbox (not just the row itself) checks it and turns the row amber
without double-toggling, and selecting a Current Styles row at the same time
shows indigo there - the two panels' selected states read as visually
distinct. `npm run build`/`test` (84/84)/`lint` all still pass.

### 2026-09-18 — "Mash it" merges straight into an already-selected New Styles
target, skipping MergeDialog

Per direct request: previously, "Mash it" (`StyleReportPanel`'s footer
button) always called `actions.openMergeDialog()` regardless of whether a
New Styles row was already picked as a merge target - so a user who'd
selected entries on the left *and* a target on the right still had to
confirm through the dialog, duplicating what `UserStylesPanel`'s own inline
"Merge N selected here" button already does with no dialog at all.

`App.tsx` now passes `StyleReportPanel` a new local `onMashIt` in place of
the inline `() => actions.openMergeDialog()`: when
`state.selectedTargetStyleId` is set (a New Styles row is picked), it calls
`actions.mergeSelectedIntoTarget()` directly - the exact same action
`UserStylesPanel`'s "Merge N selected here" button already calls, reading
`state.selectedVariantIds`/`state.selectedTargetStyleId` off `stateRef`
itself, so no new plumbing was needed - and only falls back to
`actions.openMergeDialog()` when no target is selected. "Mash it" stays
disabled until at least one Style Report entry is selected either way (see
`StyleReportPanel`'s own `disabled={selectedIds.size === 0}`), so the only
extra branch needed was on `selectedTargetStyleId`.

Verified in Chrome: with no New Styles selection, selecting a Current Styles
entry and clicking "Mash it" opens `MergeDialog` as before (confirmed then
canceled, selection preserved). With a New Styles target *also* selected
(amber-highlighted "List Bullet", from the previous entry's checkbox work),
clicking "Mash it" merged immediately with no dialog - "List Bullet"'s
occurrence count went 0×→2×, the merged Style Report entry disappeared from
Current Styles, and the progress bar updated to "1 of 19 merged", all
identical to what the inline "Merge N selected here" button already
produces. `npm run build`/`test` (84/84)/`lint` all still pass.

### 2026-09-18 — "Attach custom Word styles" rename; "+ Defaults" gets an
orange, attention-grabbing treatment

Per direct request: `AttachReferenceDocButton.tsx`'s label changed from
"Attach style file (recommended)" to **"Attach custom Word styles"** (its own
doc comment in `UserStylesPanel.tsx`, and the matching README walkthrough
step, updated to match - same "keep prose in sync with actual UI copy"
practice this file's history already follows).

`UserStylesPanel.tsx`'s **"+ Defaults"** button changed from its previous
plain bordered/ghost style (`border-indigo-200 text-slate-300`, blending into
the dark panel header) to a solid, high-contrast **orange**
(`bg-orange-600 hover:bg-orange-700 text-white`) - orange-600 against white
text clears WCAG AA's small-text contrast threshold, and reads as a
deliberate callout next to "+ New Style"'s now-comparatively-quiet outline
style, the same "solid = primary/highlighted action" language the rest of
the app already uses for e.g. "Mash it"/"Attach custom Word styles"
themselves (indigo-600 there; orange here is the one deliberate departure,
so this specific action reads as distinct from the app's indigo action
color).

A first pass also added a small "TRY ME" callout (same `text-[10px]
font-semibold uppercase` treatment as AppHeader's "BETA" tag) to the left of
the button - reverted per immediate follow-up feedback ("the orange button
is enough"). Its brief presence surfaced a real, worth-recording layout
constraint though: `UserStylesPanel`'s header row (title + count + info icon
on the left, `shrink-0` buttons on the right, in a ~340px grid column) was
already right at its width budget with just two buttons - adding even a
short new element to that row was enough to force "New Styles" to word-wrap
mid-phrase rather than degrading gracefully. Worth remembering before adding
anything else to that row: either trim from elsewhere in the same row first,
or give the row `flex-wrap` so an overflow drops the button group to a
second line instead of breaking the title text.

Verified in Chrome (both at ~1456px and ~1400px window widths, the range
where the header wrapping above was actually observed): "New Styles (0)"
renders on one line, "+ Defaults" reads clearly as solid orange against the
dark header, and "Attach custom Word styles" wraps cleanly as a two-line
button label without clipping. `npm run build`/`test` (84/84)/`lint` all
still pass.

### 2026-09-18 — Real Help modal copy, in its own editable content file

`HelpModal.tsx` had shipped since the original build with placeholder filler
copy - its own doc comment said as much ("Content is filler copy - swap in
real walkthrough/FAQ text when it's written"), and it had drifted from the
actual UI besides (referenced a "New styles" button and a "Do it" button,
neither of which exist - see the 2026-09-18 README-sync entry earlier in
this log, which fixed the same drift in README.md but explicitly left
HelpModal out of scope). Per direct request, wrote the real thing.

Per the request's own explicit requirement, the copy itself lives entirely
outside the component: new `src/content/help-content.md`, plain prose with
`#`/`##` headings and blank-line-separated paragraphs (plus one `- ` bulleted
list, in the "New Styles" section) - no JSX, no string-escaping, nothing
that isn't editable by just opening the file and changing sentences.
Content was written from this project's own actual workflow (every panel,
button and behavior it names was checked against the current UI - "Mash it"
+ "+ Defaults" + "Attach custom Word styles" + "Customise your own style
file" + the two merge paths (direct-into-a-selected-target vs. the dialog,
see the "Mash it" entry above) + Undo/Clear list/Mash a different file +
Save your file's `-RIPPED` suffix + what's out of scope (headers/footers/
footnotes, mobile) - not copied from README.md's own prose, but covering the
same ground independently since the two serve different audiences (in-app
help vs. a repo README).

`HelpModal.tsx` itself is now just dialog chrome (heading, scrollable body,
Close button) plus a small parser (`parseHelpContent`/`toBlocks`) that turns
the file's plain-text structure into rendered paragraphs/lists - imported via
Vite's `?raw` suffix (`import helpContentRaw from '../content/help-content.md?raw'`,
type-checked already via the `vite/client` types `tsconfig.app.json` already
declares, so no new type declarations were needed) and parsed once at module
load since the content is static. No test covered the old placeholder copy,
so none needed updating.

Verified in Chrome: opened the Help modal, confirmed the real title/intro/
section content renders (not the old placeholder), and scrolled to confirm
the "New Styles" section's three bullet points render as a real `<ul>`, not
run-on prose with stray "- " characters. `npm run build`/`test` (84/84)/
`lint` all still pass.

### 2026-09-18 — "Attach custom Word styles" recolored dark gray

Per direct request: `AttachReferenceDocButton.tsx`'s button changed from
indigo (`bg-indigo-600 hover:bg-indigo-700`, the app's general primary-action
color, shared with "Mash it" and others) to dark gray
(`bg-slate-600 hover:bg-slate-700`) - its own action (attaching a second
Word file as a style source) is optional/secondary next to the main merge
workflow, so a quieter color than the primary indigo actions reads as more
accurate now. `npm run build`/`test` (84/84)/`lint` all still pass; verified
in Chrome.

#### Verified

`npm run build`/`test` (84/84 - the existing `defaultStyles.test.ts`/
`workspaceReducer.test.tsx` suites needed no changes, since both were already
generic over the actual style set rather than hardcoding names/counts) /
`lint` all pass (same single pre-existing `no-constant-binary-expression`
warning as every prior pass). Re-verified end-to-end in Chrome: uploaded the
new `public/CLEAN-STYLES.docx` itself, confirmed all three categories render
with correct counts and item order (Body text styles 5/5, Heading Styles
9/9, List styles 6/6) and correct fonts (Aptos, or "default font" for the
null-fontFamily entries - no Arial anywhere in the checklist), and confirmed
"+ Defaults" produces exactly 20 records with no duplicates. The root
directory was confirmed clean afterward (`STYLE-CATEGORIES.docx` and its
Word lock file both gone).
