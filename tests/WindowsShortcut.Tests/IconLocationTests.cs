using System;
using Xunit;

namespace WindowsShortcut.Tests;

public class IconLocationTests
{
    [Theory]
    [InlineData("0")]
    [InlineData("%windir%\\system32\\shell32.dll")]
    [InlineData(",0")]
    public void Parse_ThrowFormatException(string s)
    {
        Assert.Throws<FormatException>(() => IconLocation.Parse(s));
    }

    [Theory]
    [InlineData("C:\\Windows\\System32\\shell32.dll,0", "C:\\Windows\\System32\\shell32.dll", 0)]
    [InlineData("%windir%\\system32\\user32.dll,1", "%windir%\\system32\\user32.dll", 1)]
    public void ParseIconLocation(string s, string expectedPath, int expectedIndex)
    {
        IconLocation iconLocation = IconLocation.Parse(s);

        Assert.Equal(expectedPath, iconLocation.Path);
        Assert.Equal(expectedIndex, iconLocation.Index);
    }
}
