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
