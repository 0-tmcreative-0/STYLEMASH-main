import { useState } from 'react'
import { FontAwesomeIcon } from '@fortawesome/react-fontawesome'
import { faChevronRight, faDownload } from '@fortawesome/free-solid-svg-icons'
import {
  DEFAULT_STYLE_CATEGORIES,
  groupDefaultStylesByCategory,
  type DefaultStyleCategory,
  type DefaultStyleDefinition,
} from '../lib/ooxml/defaultStyles'
import { signatureToCss } from '../lib/signatureToCss'
import { describeSignature } from '../lib/styleDescriptions'
import { FaCheckbox } from './FaCheckbox'

interface DefaultStylesChecklistProps {
  /** Names of every DEFAULT_STYLES entry currently checked - what "+
   * Defaults" will actually apply. Lives in useDocxWorkspace, independent
   * of the loaded-document state (see that hook's own note). */
  enabledNames: Set<string>
  onToggle: (name: string) => void
}

/** One row - a style definition to include or skip. Deliberately the same
 * shape as StyleVariantRow's row in the Current Styles panel (FaCheckbox +
 * live-styled sample line + a describeSignature() line underneath) rather
 * than its own layout, so ticking a box here and ticking one there read as
 * the same gesture. The two real differences from StyleVariantRow: there's
 * no origin line (a bundled default has no document to have come from) and
 * no occurrence badge (nothing's been counted yet). */
function DefaultStyleRow({
  def,
  enabledNames,
  onToggle,
}: {
  def: DefaultStyleDefinition
  enabledNames: Set<string>
  onToggle: (name: string) => void
}) {
  const checked = enabledNames.has(def.name)
  return (
    <li
      onClick={() => onToggle(def.name)}
      className={`flex cursor-pointer items-start gap-3 border-b border-l-4 border-slate-200 px-4 py-3 transition-colors last:border-b-0 ${
        checked ? 'border-l-indigo-500 bg-white hover:bg-slate-50' : 'border-l-transparent bg-slate-50/60 hover:bg-slate-50'
      }`}
    >
      <FaCheckbox
        checked={checked}
        onToggle={() => onToggle(def.name)}
        label={`Include "${def.name}" in + Defaults`}
      />
      <div className="min-w-0 flex-1">
        <p className="truncate text-base" style={signatureToCss(def.targetSignature)}>
          {def.listPreviewText && <span className="mr-1 text-slate-400">{def.listPreviewText}</span>}
          {def.name}
        </p>
        <p className="mt-1 truncate text-xs text-slate-500">{describeSignature(def.targetSignature)}</p>
      </div>
    </li>
  )
}

/** One collapsible category section (e.g. "Headings") - its own local
 * open/closed state, since sections expand/collapse independently. The
 * height transition is a CSS grid-template-rows trick (0fr -> 1fr) rather
 * than a JS-measured max-height: it animates to the content's real height
 * with no measurement, and never clips content that grows after the fact. */
function CategorySection({
  category,
  styles,
  enabledNames,
  onToggle,
}: {
  category: DefaultStyleCategory
  styles: DefaultStyleDefinition[]
  enabledNames: Set<string>
  onToggle: (name: string) => void
}) {
  const [isOpen, setIsOpen] = useState(false)
  const checkedCount = styles.filter((s) => enabledNames.has(s.name)).length

  return (
    <div className="overflow-hidden rounded-md border border-slate-200">
      <button
        type="button"
        onClick={() => setIsOpen((v) => !v)}
        aria-expanded={isOpen}
        className="flex w-full items-center justify-between gap-2 bg-white px-3 py-2 text-left text-xs font-medium text-slate-700 hover:bg-slate-50"
      >
        <span className="flex items-center gap-1.5">
          <FontAwesomeIcon
            icon={faChevronRight}
            aria-hidden="true"
            className={`text-[10px] text-slate-400 transition-transform duration-200 ${isOpen ? 'rotate-90' : ''}`}
          />
          {category}
        </span>
        <span className="text-[11px] font-normal text-slate-400">
          {checkedCount}/{styles.length}
        </span>
      </button>

      <div
        className={`grid transition-[grid-template-rows] duration-300 ease-in-out ${
          isOpen ? 'grid-rows-[1fr]' : 'grid-rows-[0fr]'
        }`}
      >
        <div className="overflow-hidden">
          <ul className="border-t border-slate-200">
            {styles.map((def) => (
              <DefaultStyleRow key={def.name} def={def} enabledNames={enabledNames} onToggle={onToggle} />
            ))}
          </ul>
        </div>
      </div>
    </div>
  )
}

/** The "Customise your own style file" content - lets a user pick which of
 * StyleMash's bundled starter styles (see defaultStyles.ts - extracted from
 * public/CLEAN-STYLES.docx) "+ Defaults" actually applies, grouped by
 * category (Body text styles, Heading Styles, List styles - see
 * DefaultStyleCategory, taken directly from CLEAN-STYLES.docx's own red
 * section headings) into collapsible sections so 20 styles don't read as one
 * long undifferentiated list. Also the one place StyleMash offers
 * CLEAN-STYLES.docx itself as a
 * download, for a user who'd rather open it in Word and build their own
 * variant from scratch than pick and choose here.
 *
 * Lives directly inside UserStylesPanel (the "New Styles" panel) - toggled
 * open/closed by AppHeader's "Customise your own style file" button - rather
 * than in a separate modal, so the grouping sits with the list it actually
 * governs. Pure content with no chrome of its own (no heading, no
 * open/close control) - the panel embedding this owns both. */
export function DefaultStylesChecklist({ enabledNames, onToggle }: DefaultStylesChecklistProps) {
  const groups = groupDefaultStylesByCategory()

  return (
    <div className="space-y-2">
      <p className="text-xs text-slate-500">
        Choose which bundled starter styles "+ Defaults" populates below. Unchecked styles are simply
        skipped.
      </p>

      <a
        href="/CLEAN-STYLES.docx"
        download
        className="inline-flex w-fit items-center gap-1.5 rounded-md border border-indigo-200 px-2 py-1 text-xs font-medium text-indigo-700 hover:bg-indigo-50"
      >
        <FontAwesomeIcon icon={faDownload} aria-hidden="true" />
        Download reference style file
      </a>

      <div className="max-h-72 space-y-1.5 overflow-y-auto pr-1">
        {DEFAULT_STYLE_CATEGORIES.map((category) => (
          <CategorySection
            key={category}
            category={category}
            styles={groups.get(category) ?? []}
            enabledNames={enabledNames}
            onToggle={onToggle}
          />
        ))}
      </div>
    </div>
  )
}
