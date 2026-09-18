interface SaveButtonProps {
  disabled: boolean
  isSaving: boolean
  onSave: () => void
  className?: string
}

/** Always-available save action. Downloads the (possibly edited) document
 * locally in its original format, filename suffixed with "-RIPPED" - never
 * uploads anything anywhere, since StyleMash does all processing
 * client-side in the browser. Rendered in StyleReportPanel's footer, below
 * the Undo/"Mash it" row. */
export function SaveButton({ disabled, isSaving, onSave, className = '' }: SaveButtonProps) {
  return (
    <button
      type="button"
      disabled={disabled || isSaving}
      onClick={onSave}
      className={`rounded-md bg-emerald-600 px-3 py-1.5 text-xs font-medium text-white enabled:hover:bg-emerald-700 disabled:cursor-not-allowed disabled:bg-slate-300 ${className}`}
    >
      {isSaving ? 'Saving…' : 'Save your file'}
    </button>
  )
}
