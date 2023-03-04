using System;
using System.ComponentModel;
using Xunit;

namespace WindowsShortcut.Tests;

public class HotKeyTests
{
    [Theory]
    [InlineData(-1)]
    [InlineData(0x10000)]
    public void Construct_ThrowArgumentOutOfRangeException(int key)
    {
        Assert.Throws<ArgumentOutOfRangeException>(() => new HotKey(key));
    }

    [Theory]
    [InlineData(0x0700)]
    [InlineData(0x0001)]
    [InlineData(0x0092)]
    public void Construct_ThrowArgumentException(int key)
    {
        Assert.Throws<ArgumentException>(() => new HotKey(key));
    }

    [Theory]
    [InlineData(-1)]
    [InlineData(0x07)]
    public void SetModifierKeys_ThrowArgumentOutOfRangeException(int mKey)
    {
        HotKey hotKey = new();

        Assert.Throws<ArgumentOutOfRangeException>(() => hotKey.ModifierKeys = (ModifierKeys)mKey);
    }

    [Theory]
    [InlineData(-1)]
    [InlineData(0x01)]
    [InlineData(0x92)]
    public void SetVirtualKey_ThrowInvalidEnumArgumentException(int vKey)
    {
        HotKey hotKey = new();

        Assert.Throws<InvalidEnumArgumentException>(() => hotKey.VirtualKey = (VirtualKey)vKey);
    }

    [Theory]
    [InlineData(0, ModifierKeys.None, VirtualKey.None)]
    [InlineData(0x0330, ModifierKeys.Shift | ModifierKeys.Control, VirtualKey.Number0)]
    public void ConstructHotKey(int key, ModifierKeys expectedMKey, VirtualKey expectedVKey)
    {
        HotKey hotKey = new(key);

        Assert.Equal(expectedMKey, hotKey.ModifierKeys);
        Assert.Equal(expectedVKey, hotKey.VirtualKey);
    }

    [Theory]
    [InlineData(0, ModifierKeys.None, VirtualKey.None)]
    [InlineData(0x0691, ModifierKeys.Control | ModifierKeys.Alt, VirtualKey.ScrollLock)]
    public void SetKeys(int expected, ModifierKeys mKey, VirtualKey vKey)
    {
        HotKey hotKey = new()
        {
            ModifierKeys = mKey,
            VirtualKey = vKey
        };

        Assert.Equal(expected, hotKey.RawData);
    }
}
