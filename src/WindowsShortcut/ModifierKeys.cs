using System;

namespace WindowsShortcut;

/// <summary>
/// Specifies modifier keys to use as a keyboard shortcut for a shell link.
/// </summary>
[Flags]
public enum ModifierKeys
{
    None = 0x00,
    Shift = 0x01,
    Control = 0x02,
    Alt = 0x04
}
