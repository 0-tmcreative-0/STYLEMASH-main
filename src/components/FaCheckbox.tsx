import { FontAwesomeIcon } from '@fortawesome/react-fontawesome'
import { faSquare as faSquareOutline } from '@fortawesome/free-regular-svg-icons'
import { faSquareCheck } from '@fortawesome/free-solid-svg-icons'

interface FaCheckboxProps {
  checked: boolean
  onToggle: () => void
  label: string
  className?: string
}

/** A checkbox rendered as a FontAwesome glyph (checked/unchecked square)
 * instead of a native checkbox's browser-drawn box, so it reads consistently
 * with the rest of StyleMash's UI rather than switching styles per platform.
 * Used both by the "Customise your own style file" panel's per-style
 * checkboxes and by StyleVariantRow's row checkbox - one visual so ticking a
 * box means the same thing (and looks the same) whether you're picking
 * which defaults to bring in or selecting Style Report entries to merge.
 *
 * A real, if visually hidden, `<input type="checkbox">` still backs the
 * icon - same "sr-only" pattern StyleVariantRow used before this component
 * existed - so screen readers and keyboard navigation (Tab + Space) keep
 * working; the icon is purely decorative (aria-hidden) on top of it. */
export function FaCheckbox({ checked, onToggle, label, className = '' }: FaCheckboxProps) {
  return (
    <span className={`relative inline-flex shrink-0 items-center ${className}`}>
      <input
        type="checkbox"
        checked={checked}
        onChange={onToggle}
        onClick={(e) => e.stopPropagation()}
        className="sr-only"
        aria-label={label}
      />
      <FontAwesomeIcon
        // The regular-style (outline) square for "unchecked" and the
        // solid checkmark-in-a-square for "checked" - free-solid-svg-icons
        // has no outline square, and a *filled* gray square read as a solid
        // blob rather than an empty box, so the unchecked state borrows
        // from free-regular-svg-icons instead.
        icon={checked ? faSquareCheck : faSquareOutline}
        aria-hidden="true"
        className={`pointer-events-none text-lg transition-all duration-150 ${
          checked ? 'scale-110 text-indigo-600' : 'scale-100 text-slate-400'
        }`}
      />
    </span>
  )
}
