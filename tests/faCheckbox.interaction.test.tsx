import { act } from 'react'
import { createRoot } from 'react-dom/client'
import { describe, expect, it, vi } from 'vitest'
import { FaCheckbox } from '../src/components/FaCheckbox'

// Same flag (and reason) as styleVariantRow.interaction.test.tsx.
;(globalThis as { IS_REACT_ACT_ENVIRONMENT?: boolean }).IS_REACT_ACT_ENVIRONMENT = true

function renderCheckbox(checked: boolean) {
  const container = document.createElement('div')
  document.body.appendChild(container)
  const root = createRoot(container)
  const onToggle = vi.fn()

  act(() => {
    root.render(<FaCheckbox checked={checked} onToggle={onToggle} label="Include this style" />)
  })

  return { container, onToggle }
}

describe('FaCheckbox', () => {
  it('renders a real checkbox input backing the icon, for keyboard/screen-reader use', () => {
    const { container } = renderCheckbox(false)
    const input = container.querySelector('input[type="checkbox"]') as HTMLInputElement
    expect(input).not.toBeNull()
    expect(input.checked).toBe(false)
    expect(input.className).toContain('sr-only')
  })

  it('clicking the input calls onToggle exactly once', () => {
    const { container, onToggle } = renderCheckbox(false)
    const input = container.querySelector('input[type="checkbox"]')!

    act(() => {
      input.dispatchEvent(new MouseEvent('click', { bubbles: true }))
    })

    expect(onToggle).toHaveBeenCalledTimes(1)
  })

  it('shows a checked icon when checked=true and an unchecked one when false', () => {
    const checkedRender = renderCheckbox(true)
    const checkedIcon = checkedRender.container.querySelector('svg')!
    expect(checkedIcon.getAttribute('data-icon')).toBe('square-check')

    const uncheckedRender = renderCheckbox(false)
    const uncheckedIcon = uncheckedRender.container.querySelector('svg')!
    expect(uncheckedIcon.getAttribute('data-icon')).toBe('square')
  })
})
