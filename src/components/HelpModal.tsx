import helpContentRaw from '../content/help-content.md?raw'

interface HelpModalProps {
  onClose: () => void
}

/** One chunk of a section's body: an ordinary paragraph, or a "- " bulleted
 * list (the one bit of structure beyond plain prose the content file uses -
 * see help-content.md's "New Styles" section). */
type HelpBlock = { kind: 'paragraph'; text: string } | { kind: 'list'; items: string[] }

interface HelpSection {
  heading: string
  blocks: HelpBlock[]
}

/** Splits one heading's worth of raw text (already isolated by the caller)
 * into blocks on blank lines - a chunk where every line starts with "- "
 * becomes a list block, anything else becomes a paragraph (its own line
 * breaks collapsed to spaces, since the content file wraps prose at
 * whatever width is convenient to edit, not for display). */
function toBlocks(text: string): HelpBlock[] {
  return text
    .trim()
    .split(/\n{2,}/)
    .map((chunk) => chunk.trim())
    .filter(Boolean)
    .map((chunk): HelpBlock => {
      const lines = chunk.split('\n').map((l) => l.trim())
      if (lines.every((l) => l.startsWith('- '))) {
        return { kind: 'list', items: lines.map((l) => l.slice(2).trim()) }
      }
      return { kind: 'paragraph', text: lines.join(' ') }
    })
}

/** Parses help-content.md's plain "# title" / "## heading" / blank-line
 * paragraph (plus the occasional "- " list) structure into data this
 * component can render - the whole point of keeping the copy in its own
 * file (see that file's own intent) is that anyone editing StyleMash's help
 * text only ever touches plain prose, never this component's JSX. */
function parseHelpContent(raw: string): { title: string; intro: HelpBlock[]; sections: HelpSection[] } {
  const [introBlock, ...sectionBlocks] = raw.trim().split(/\n(?=## )/)
  const introLines = introBlock.split('\n')
  const title = (introLines[0] ?? '').replace(/^#\s*/, '').trim()
  const intro = toBlocks(introLines.slice(1).join('\n'))

  const sections = sectionBlocks.map((block) => {
    const lines = block.split('\n')
    const heading = (lines[0] ?? '').replace(/^##\s*/, '').trim()
    return { heading, blocks: toBlocks(lines.slice(1).join('\n')) }
  })

  return { title, intro, sections }
}

function HelpBlocks({ blocks }: { blocks: HelpBlock[] }) {
  return (
    <>
      {blocks.map((block, i) =>
        block.kind === 'list' ? (
          <ul key={i} className="list-disc space-y-1 pl-5">
            {block.items.map((item, j) => (
              <li key={j}>{item}</li>
            ))}
          </ul>
        ) : (
          <p key={i}>{block.text}</p>
        ),
      )}
    </>
  )
}

const { title, intro, sections } = parseHelpContent(helpContentRaw)

/** Help modal, opened from AppHeader's "Help" button. All the actual
 * walkthrough copy lives in src/content/help-content.md, not here - this
 * component only supplies the dialog chrome (heading, scroll area, Close
 * button) and a small parser (see parseHelpContent above) that turns that
 * file's plain "# / ## / blank-line-paragraph" structure into rendered
 * sections. Parsed once at module load, not per-render, since the content
 * is static. */
export function HelpModal({ onClose }: HelpModalProps) {
  return (
    <div className="fixed inset-0 z-50 flex items-center justify-center bg-black/40 px-4">
      <div className="w-full max-w-md rounded-lg bg-white p-5 shadow-xl">
        <h2 className="text-3xl font-semibold text-slate-800">Help</h2>
        <div className="mt-3 max-h-80 space-y-3 overflow-y-auto text-sm text-slate-600">
          <h3 className="text-2xl font-semibold text-slate-800">{title}</h3>
          <HelpBlocks blocks={intro} />
          {sections.map((section) => (
            <div key={section.heading} className="space-y-2">
              <h4 className="text-lg font-semibold text-slate-800">{section.heading}</h4>
              <HelpBlocks blocks={section.blocks} />
            </div>
          ))}
        </div>
        <div className="mt-4 flex justify-end">
          <button
            type="button"
            onClick={onClose}
            className="rounded-md bg-indigo-600 px-3 py-1.5 text-xs font-medium text-white hover:bg-indigo-700"
          >
            Close
          </button>
        </div>
      </div>
    </div>
  )
}
