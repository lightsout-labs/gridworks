import { describe, expect, it } from 'vitest';
import type { KeyPress } from '../grid/navigation';
import { shortcut } from './shortcuts';

const press = (key: string, modifiers: Partial<KeyPress> = {}): KeyPress => ({
  key,
  shift: false,
  ctrl: false,
  meta: false,
  alt: false,
  ...modifiers,
});

describe('shortcut', () => {
  it('undoes with Ctrl+Z and redoes with Ctrl+Y', () => {
    expect(shortcut(press('z', { ctrl: true }), false)).toBe('undo');
    expect(shortcut(press('y', { ctrl: true }), false)).toBe('redo');
    // With Caps Lock on.
    expect(shortcut(press('Z', { ctrl: true }), false)).toBe('undo');
  });

  it('copies with Ctrl+C, cuts with Ctrl+X and pastes with Ctrl+V', () => {
    expect(shortcut(press('c', { ctrl: true }), false)).toBe('copy');
    expect(shortcut(press('x', { ctrl: true }), false)).toBe('cut');
    expect(shortcut(press('v', { ctrl: true }), false)).toBe('paste');
    expect(shortcut(press('v', { meta: true }), true)).toBe('paste');
  });

  it('uses Cmd in place of Ctrl on a Mac', () => {
    expect(shortcut(press('z', { meta: true }), true)).toBe('undo');
    expect(shortcut(press('y', { meta: true }), true)).toBe('redo');
    expect(shortcut(press('z', { ctrl: true }), true)).toBeNull();
    expect(shortcut(press('z', { meta: true }), false)).toBeNull();
  });

  it('ignores other keys and other modifiers', () => {
    expect(shortcut(press('z'), false)).toBeNull();
    expect(shortcut(press('q', { ctrl: true }), false)).toBeNull();
    expect(shortcut(press('z', { ctrl: true, alt: true }), false)).toBeNull();
    expect(shortcut(press('y', { ctrl: true, shift: true }), false)).toBeNull();
  });
});
