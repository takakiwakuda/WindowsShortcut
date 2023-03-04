using System;
using System.ComponentModel;

namespace WindowsShortcut;

/// <summary>
/// Specifies a keyboard shortcut for a shell link.
/// </summary>
public sealed class HotKey
{
    /// <summary>
    /// Not specified
    /// </summary>
    public static HotKey None => new();

    /// <summary>
    /// Control + Alt
    /// </summary>
    private const int MaxModifierKeys = 0x06;

    /// <summary>
    /// Not specified
    /// </summary>
    private const int MinModifierKeys = 0x00;

    /// <summary>
    /// Gets or sets the modifier keys of the hotkey.
    /// </summary>
    public ModifierKeys ModifierKeys
    {
        get => (ModifierKeys)_modifierKeys;
        set
        {
            int mKey = (int)value;
            if (mKey < MinModifierKeys || mKey > MaxModifierKeys)
            {
                throw new ArgumentOutOfRangeException(nameof(value));
            }

            SetRawData(mKey, _virtualKey);
            _modifierKeys = mKey;
        }
    }

    /// <summary>
    /// Gets or sets the virtual key of the hotkey.
    /// </summary>
    public VirtualKey VirtualKey
    {
        get => (VirtualKey)_virtualKey;
        set
        {
            int vKey = (int)value;
            if (!Enum.IsDefined(typeof(VirtualKey), value))
            {
                throw new InvalidEnumArgumentException(nameof(value), vKey, typeof(VirtualKey));
            }

            SetRawData(_modifierKeys, vKey);
            _virtualKey = vKey;
        }
    }

    internal ushort RawData => _rawData;

    private int _modifierKeys;
    private int _virtualKey;
    private ushort _rawData;

    /// <summary>
    /// Initializes a new instance of the <see cref="HotKey"/> class.
    /// </summary>
    public HotKey()
    {
    }

    /// <summary>
    /// Initializes a new instance of the <see cref="HotKey"/> class with the specified key.
    /// </summary>
    /// <param name="key">A number that represents the keyboard shortcut.</param>
    /// <exception cref="ArgumentOutOfRangeException">
    /// <paramref name="key"/> is less than <see cref="ushort.MinValue"/> or greater than <see cref="ushort.MaxValue"/>.
    /// </exception>
    /// <exception cref="ArgumentException">
    /// <para>
    /// High-order byte of <paramref name="key"/> is not a valid modifier key.
    /// </para>
    /// -or-
    /// <para>
    /// Low-order byte of <paramref name="key"/> is not a valid virtual key.
    /// </para>
    /// </exception>
    public HotKey(int key)
    {
        if (key < ushort.MinValue || key > ushort.MaxValue)
        {
            throw new ArgumentOutOfRangeException(nameof(key));
        }

        SetKeys(key);
    }

    /// <summary>
    /// Returns a string representing the current hotkey.
    /// </summary>
    /// <returns>A string representing the current hotkey.</returns>
    public override string ToString()
    {
        if (RawData == 0)
        {
            return "None";
        }

        string s = ModifierKeys.ToString().Replace(", ", "+");
        return s + "+" + VirtualKey.ToString();
    }

    private void SetKeys(int key)
    {
        if (key == 0)
        {
            return;
        }

        int mKey = key >> 8;
        if (mKey > MaxModifierKeys)
        {
            string message = $"High-order byte of the value 0x{key:X4} is not a valid modifier key.";
            throw new ArgumentException(message, nameof(key));
        }

        int vKey = key & 0x00FF;
        if (!Enum.IsDefined(typeof(VirtualKey), vKey))
        {
            string message = $"Low-order byte of the value 0x{key:X4} is not a valid virtual key.";
            throw new ArgumentException(message, nameof(key));
        }

        _modifierKeys = mKey;
        _virtualKey = vKey;
        _rawData = (ushort)key;
    }

    private void SetRawData(int mKey, int vKey) => _rawData = (ushort)((mKey << 8) + vKey);
}
