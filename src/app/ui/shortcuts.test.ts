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

  it('fills down with Ctrl+D and right with Ctrl+R', () => {
    expect(shortcut(press('d', { ctrl: true }), false)).toBe('fillDown');
    expect(shortcut(press('r', { ctrl: true }), false)).toBe('fillRight');
    expect(shortcut(press('d', { meta: true }), true)).toBe('fillDown');
  });

  it('toggles bold, italic and underline with Ctrl+B, I and U', () => {
    expect(shortcut(press('b', { ctrl: true }), false)).toBe('bold');
    expect(shortcut(press('i', { ctrl: true }), false)).toBe('italic');
    expect(shortcut(press('u', { meta: true }), true)).toBe('underline');
  });

  it('finds with Ctrl+F and replaces with Ctrl+H', () => {
    expect(shortcut(press('f', { ctrl: true }), false)).toBe('find');
    expect(shortcut(press('h', { ctrl: true }), false)).toBe('replace');
    expect(shortcut(press('f', { meta: true }), true)).toBe('find');
    // macOS hides the browser on Cmd+H.
    expect(shortcut(press('h', { ctrl: true }), true)).toBe('replace');
  });

  it('prints with Ctrl+P', () => {
    expect(shortcut(press('p', { ctrl: true }), false)).toBe('print');
    expect(shortcut(press('p', { meta: true }), true)).toBe('print');
  });

  it('leaves Ctrl+C, X and V to the clipboard events', () => {
    expect(shortcut(press('c', { ctrl: true }), false)).toBeNull();
    expect(shortcut(press('x', { ctrl: true }), false)).toBeNull();
    expect(shortcut(press('v', { meta: true }), true)).toBeNull();
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
