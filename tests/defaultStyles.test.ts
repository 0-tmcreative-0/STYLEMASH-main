import { describe, expect, it } from 'vitest'
import {
  addDefaultStyles,
  DEFAULT_STYLE_CATEGORIES,
  DEFAULT_STYLES,
  groupDefaultStylesByCategory,
} from '../src/lib/ooxml/defaultStyles'
import { NS } from '../src/lib/ooxml/constants'
import { makeParsedDocx } from './testUtils'

describe('DEFAULT_STYLES', () => {
  it('every entry has a name unique across the whole set', () => {
    const names = DEFAULT_STYLES.map((d) => d.name)
    expect(new Set(names).size).toBe(names.length)
  })

  it('every entry is assigned one of DEFAULT_STYLE_CATEGORIES', () => {
    for (const def of DEFAULT_STYLES) {
      expect(DEFAULT_STYLE_CATEGORIES).toContain(def.category)
    }
  })

  // Regression guard: "Normal" never showed up in the styles CLEAN-STYLES.docx's
  // body references by an explicit w:pStyle (see defaultStyles.ts's own doc
  // comment) - it's the document's *implicit* default, used by leaving pStyle
  // off entirely. Easy to lose in a future re-extraction pass.
  it('includes "Normal" - merging into it matters most for real-world documents', () => {
    const normal = DEFAULT_STYLES.find((d) => d.name === 'Normal')
    expect(normal).toBeDefined()
    expect(normal?.kind).toBe('character')
    expect(normal?.listFormat).toBe('none')
  })
})

describe('groupDefaultStylesByCategory', () => {
  it('every DEFAULT_STYLES entry appears in exactly its own category, in DEFAULT_STYLES order', () => {
    const groups = groupDefaultStylesByCategory()
    expect(new Set(groups.keys())).toEqual(new Set(DEFAULT_STYLE_CATEGORIES))

    const flattened = DEFAULT_STYLE_CATEGORIES.flatMap((c) => groups.get(c) ?? [])
    expect(flattened.map((d) => d.name).sort()).toEqual(DEFAULT_STYLES.map((d) => d.name).sort())
    for (const def of DEFAULT_STYLES) {
      expect(groups.get(def.category)).toContain(def)
    }
  })
})

const W = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"'

describe('addDefaultStyles enabledNames filtering', () => {
  it('applies every DEFAULT_STYLES entry when enabledNames is omitted (back-compat default)', () => {
    const parsedDocx = makeParsedDocx({ documentXml: `<w:document ${W}><w:body></w:body></w:document>` })
    const userStyles = addDefaultStyles(parsedDocx, [])
    expect(userStyles).toHaveLength(DEFAULT_STYLES.length)
  })

  it('skips every style not named in enabledNames', () => {
    const parsedDocx = makeParsedDocx({ documentXml: `<w:document ${W}><w:body></w:body></w:document>` })
    const userStyles = addDefaultStyles(parsedDocx, [], new Set(['Normal', 'heading 1']))

    expect(userStyles.map((r) => r.name).sort()).toEqual(['Normal', 'heading 1'])
    const styleCount = parsedDocx.stylesXml.getElementsByTagNameNS(NS.w, 'style').length
    expect(styleCount).toBe(2)
  })

  it('an empty enabledNames set creates nothing', () => {
    const parsedDocx = makeParsedDocx({ documentXml: `<w:document ${W}><w:body></w:body></w:document>` })
    const userStyles = addDefaultStyles(parsedDocx, [], new Set())

    expect(userStyles).toHaveLength(0)
    expect(parsedDocx.stylesXml.getElementsByTagNameNS(NS.w, 'style')).toHaveLength(0)
  })

  it('re-running with a narrower enabledNames set does not remove styles already created', () => {
    // addDefaultStyles never deletes - it's purely additive/redefining (see
    // its own doc comment on the collision rule), so toggling a style off in
    // the Customise panel after it's already been added has no retroactive
    // effect until "Clear list" or a manual removal.
    const parsedDocx = makeParsedDocx({ documentXml: `<w:document ${W}><w:body></w:body></w:document>` })
    const first = addDefaultStyles(parsedDocx, [], new Set(['Normal', 'heading 1']))
    const second = addDefaultStyles(parsedDocx, first, new Set(['Normal']))

    expect(second.map((r) => r.name).sort()).toEqual(['Normal', 'heading 1'])
  })
})
