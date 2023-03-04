using System;

namespace WindowsShortcut;

/// <summary>
/// Specifies an icon location for a shell link file.
/// </summary>
public sealed class IconLocation
{
    /// <summary>
    /// Gets the full path to the file that contains the icons.
    /// </summary>
    public string Path => _path;

    /// <summary>
    /// Gets an index that points to an icon in the file.
    /// </summary>
    public int Index => _index;

    private readonly string _path;
    private readonly int _index;

    /// <summary>
    /// Initializes a new instance of the <see cref="IconLocation"/> class with the specified path and index.
    /// </summary>
    /// <param name="path">The full path containing the icons.</param>
    /// <param name="index">An index pointing to an icon in <paramref name="path"/>.</param>
    /// <exception cref="ArgumentNullException">
    /// <paramref name="path"/> is <see langword="null"/>.
    /// </exception>
    /// <exception cref="ArgumentException">
    /// <paramref name="path"/> is empty.
    /// </exception>
    public IconLocation(string path, int index)
    {
        if (path is null)
        {
            throw new ArgumentNullException(nameof(path));
        }

        if (path.Length == 0)
        {
            throw new ArgumentException("The value cannot be an empty string.", nameof(path));
        }

        _path = path;
        _index = index;
    }

    /// <summary>
    /// Creates a new <see cref="IconLocation"/> using the specified string.
    /// </summary>
    /// <param name="iconLocation">A string pointing to the icon location.</param>
    /// <returns></returns>
    /// <exception cref="ArgumentNullException">
    /// <paramref name="iconLocation"/> is <see langword="null"/>.
    /// </exception>
    /// <exception cref="ArgumentException">
    /// <paramref name="iconLocation"/> is empty.
    /// </exception>
    /// <exception cref="FormatException">
    /// <paramref name="iconLocation"/> is in an invalid format.
    /// </exception>
    public static IconLocation Parse(string iconLocation)
    {
        if (iconLocation is null)
        {
            throw new ArgumentNullException(nameof(iconLocation));
        }

        if (iconLocation.Length == 0)
        {
            throw new ArgumentException("The value cannot be an empty string.", nameof(iconLocation));
        }

        ReadOnlySpan<char> span = iconLocation.AsSpan();
        int comma = span.IndexOf(',');

#if NET7_0_OR_GREATER
        if (comma < 2 || !int.TryParse(span.Slice(comma + 1), out int index))
#else
        if (comma < 2 || !int.TryParse(span.Slice(comma + 1).ToString(), out int index))
#endif
        {
            throw new FormatException($"The value '{iconLocation}' is in an invalid format.");
        }

        return new IconLocation(span.Slice(0, comma).ToString(), index);
    }

    /// <summary>
    /// Returns a string representing the current icon location.
    /// </summary>
    /// <returns>A string representing the current icon location.</returns>
    public override string ToString() => $"{Path},{Index}";
}
