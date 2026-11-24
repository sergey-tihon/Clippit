using System;
using System.Buffers.Binary;

namespace Clippit.PowerPoint.Fluent;

/// <summary>
/// Embedded OpenType (EOT) header of an embedded font part ("application/x-fontdata").
/// </summary>
internal static class EmbeddedFontHeader
{
    private const int EotSizeOffset = 0;
    private const int FontDataSizeOffset = 4;
    private const int VersionOffset = 8;
    private const int MagicNumberOffset = 34;
    private const int PrefixSize = MagicNumberOffset + sizeof(ushort);
    private const int FamilyNameSizeOffset = 82;
    private const int NameCount = 4;
    private const ushort MagicNumber = 0x504C;
    private const uint Version10 = 0x00010000;
    private const uint Version22 = 0x00020002;

    // The fields that an EOT 2.2 header has after FullName, for a font with no root string, no
    // signature and no EUDC font: Padding5 (2), RootStringSize (2), RootStringCheckSum (4),
    // EUDCCodePage (4), Padding6 (2), SignatureSize (2), EUDCFlags (4), EUDCFontSize (4).
    // The values other than zero are the ones that PowerPoint writes.
    private const int Version22TailSize = 24;
    private const int RootStringCheckSumOffset = 4;
    private const int EudcCodePageOffset = 8;
    private const uint EmptyRootStringCheckSum = 0x50475342;
    private const uint EudcCodePage = 10000;

    /// <summary>
    /// Reads a font and, when it has the EOT 1.0 header, returns it with an EOT 2.2 header.
    /// A font with another header is not read past the first bytes.
    /// </summary>
    /// <returns>Null when the font is not a well-formed EOT 1.0 file.</returns>
    public static byte[]? ReadUpgraded(Stream source)
    {
        Span<byte> prefix = stackalloc byte[PrefixSize];
        if (source.ReadAtLeast(prefix, prefix.Length, false) < prefix.Length || !IsVersion10(prefix))
            return null;

        using var buffer = new MemoryStream();
        buffer.Write(prefix);
        source.CopyTo(buffer);
        return UpgradeToVersion22(buffer.GetBuffer().AsSpan(0, (int)buffer.Length));
    }

    /// <summary>
    /// Returns an EOT 1.0 file with an EOT 2.2 header. The font data stays the same.
    /// PowerPoint for the web and the SharePoint preview do not open a presentation that has
    /// an embedded font with the EOT 1.0 header (PPTCannotInstallFonts).
    /// </summary>
    /// <returns>Null when the input is not a well-formed EOT 1.0 file.</returns>
    public static byte[]? UpgradeToVersion22(ReadOnlySpan<byte> eot)
    {
        if (!IsVersion10(eot))
            return null;

        var eotSize = BinaryPrimitives.ReadUInt32LittleEndian(eot[EotSizeOffset..]);
        var fontDataSize = BinaryPrimitives.ReadUInt32LittleEndian(eot[FontDataSizeOffset..]);
        if (eotSize != (uint)eot.Length || fontDataSize > eotSize)
            return null;

        var headerSize = (int)(eotSize - fontDataSize);
        if (!EndsAfterFullName(eot[..headerSize]))
            return null;

        var upgraded = new byte[eot.Length + Version22TailSize];
        eot[..headerSize].CopyTo(upgraded);
        eot[headerSize..].CopyTo(upgraded.AsSpan(headerSize + Version22TailSize));
        BinaryPrimitives.WriteUInt32LittleEndian(upgraded.AsSpan(EotSizeOffset), (uint)upgraded.Length);
        BinaryPrimitives.WriteUInt32LittleEndian(upgraded.AsSpan(VersionOffset), Version22);

        var tail = upgraded.AsSpan(headerSize, Version22TailSize);
        BinaryPrimitives.WriteUInt32LittleEndian(tail[RootStringCheckSumOffset..], EmptyRootStringCheckSum);
        BinaryPrimitives.WriteUInt32LittleEndian(tail[EudcCodePageOffset..], EudcCodePage);
        return upgraded;
    }

    private static bool IsVersion10(ReadOnlySpan<byte> eot) =>
        eot.Length >= PrefixSize
        && BinaryPrimitives.ReadUInt32LittleEndian(eot[VersionOffset..]) == Version10
        && BinaryPrimitives.ReadUInt16LittleEndian(eot[MagicNumberOffset..]) == MagicNumber;

    // FamilyName, StyleName, VersionName and FullName: each has 2 bytes of padding, a 2-byte
    // size and the string. The EOT 1.0 header ends after FullName.
    private static bool EndsAfterFullName(ReadOnlySpan<byte> header)
    {
        var position = FamilyNameSizeOffset - sizeof(ushort);
        for (var name = 0; name < NameCount; name++)
        {
            position += sizeof(ushort);
            if (position + sizeof(ushort) > header.Length)
                return false;
            position += sizeof(ushort) + BinaryPrimitives.ReadUInt16LittleEndian(header[position..]);
        }
        return position == header.Length;
    }
}
