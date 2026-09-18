import type { CSSProperties } from 'react'
import type { FormattingSignature } from '../types/ooxml'

/** Cheap, high-value live preview: renders a resolved FormattingSignature as
 * inline CSS so every sample line in the app - Style Report rows, User/New
 * Styles rows, the Customise checklist, the Merge Dialog's live preview, and
 * the Document Preview itself - shows roughly what the text actually looks
 * like, without needing a real Word rendering engine.
 *
 * `fontFamily` layers standard web-safe fallbacks after the signature's own
 * font name rather than rendering it verbatim: most custom/corporate fonts
 * (Aptos included - StyleMash's own bundled default font, unlikely to be
 * installed outside Microsoft 365) aren't installed in a browser, so it
 * would silently fall back to whatever generic font the browser happens to
 * pick - inconsistently between call sites, and inconsistently with the
 * Document Preview specifically, which used to apply this same fallback
 * chain on top of signatureToCss() locally rather than everyone getting it
 * for free. Every sample line in the app now renders through the identical
 * font stack, so a style's look reads the same wherever it's shown. */
export function signatureToCss(sig: FormattingSignature): CSSProperties {
  const decorations: string[] = []
  if (sig.underline) decorations.push('underline')
  if (sig.strike) decorations.push('line-through')

  const family = sig.fontFamily?.trim()

  return {
    fontFamily: family
      ? `"${family}", Calibri, "Segoe UI", Arial, sans-serif`
      : 'Calibri, "Segoe UI", Arial, sans-serif',
    fontSize: sig.fontSizeHalfPt ? `${sig.fontSizeHalfPt / 2}pt` : undefined,
    color: sig.colorValue === 'auto' ? undefined : `#${sig.colorValue}`,
    fontWeight: sig.bold ? 'bold' : 'normal',
    fontStyle: sig.italic ? 'italic' : 'normal',
    textDecorationLine: decorations.length > 0 ? decorations.join(' ') : 'none',
  }
}
