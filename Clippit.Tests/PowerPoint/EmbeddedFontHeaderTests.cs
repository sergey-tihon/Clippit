using System.Buffers.Binary;
using System.Text;
using Clippit.PowerPoint.Fluent;

namespace Clippit.Tests.PowerPoint;

public class EmbeddedFontHeaderTests
{
    private const int VersionOffset = 8;
    private const int Version22TailSize = 24;

    internal static readonly byte[] FontData = [0x00, 0x01, 0x00, 0x00, 0xDE, 0xAD, 0xBE, 0xEF, 0x01, 0x02];

    [Test]
    public async Task Version10_IsWrittenWithVersion22Header()
    {
        var source = CreateEot(0x00010000, FontData);
        var headerSize = source.Length - FontData.Length;

        var upgraded = EmbeddedFontHeader.UpgradeToVersion22(source)!;

        await Assert.That(upgraded[^FontData.Length..]).IsEquivalentTo(FontData);

        var expectedHeader = source[..headerSize];
        BinaryPrimitives.WriteUInt32LittleEndian(expectedHeader, (uint)upgraded.Length);
        BinaryPrimitives.WriteUInt32LittleEndian(expectedHeader.AsSpan(VersionOffset), 0x00020002);
        await Assert.That(upgraded[..headerSize]).IsEquivalentTo(expectedHeader);

        byte[] expectedTail = [.. new byte[4], 0x42, 0x53, 0x47, 0x50, 0x10, 0x27, .. new byte[14]];
        await Assert.That(upgraded[headerSize..^FontData.Length]).IsEquivalentTo(expectedTail);
    }

    [Test]
    public async Task Version22_IsNotChanged()
    {
        var source = CreateEot(0x00020002, FontData, new byte[Version22TailSize]);

        await Assert.That(EmbeddedFontHeader.UpgradeToVersion22(source)).IsNull();
    }

    [Test]
    public async Task Version10_WithBytesAfterFullName_IsNotChanged()
    {
        var source = CreateEot(0x00010000, FontData, [0x00, 0x00, 0x00, 0x00]);

        await Assert.That(EmbeddedFontHeader.UpgradeToVersion22(source)).IsNull();
    }

    [Test]
    public async Task Version10_WithWrongTotalSize_IsNotChanged()
    {
        var source = CreateEot(0x00010000, FontData);
        BinaryPrimitives.WriteUInt32LittleEndian(source, (uint)source.Length + 1);

        await Assert.That(EmbeddedFontHeader.UpgradeToVersion22(source)).IsNull();
    }

    [Test]
    public async Task Version10_WithNameSizePastTheHeader_IsNotChanged()
    {
        var source = CreateEot(0x00010000, FontData);
        BinaryPrimitives.WriteUInt16LittleEndian(source.AsSpan(82), ushort.MaxValue);

        await Assert.That(EmbeddedFontHeader.UpgradeToVersion22(source)).IsNull();
    }

    [Test]
    public async Task ReadUpgraded_ReadsVersion10AndLeavesVersion22()
    {
        var eot10 = CreateEot(0x00010000, FontData);
        var eot22 = CreateEot(0x00020002, FontData, new byte[Version22TailSize]);

        using var stream10 = new MemoryStream(eot10);
        using var stream22 = new MemoryStream(eot22);
        using var empty = new MemoryStream();

        await Assert
            .That(EmbeddedFontHeader.ReadUpgraded(stream10)!)
            .IsEquivalentTo(EmbeddedFontHeader.UpgradeToVersion22(eot10)!);
        await Assert.That(EmbeddedFontHeader.ReadUpgraded(stream22)).IsNull();
        await Assert.That(EmbeddedFontHeader.ReadUpgraded(empty)).IsNull();
    }

    [Test]
    public async Task NotAnEotFile_IsNotChanged()
    {
        await Assert.That(EmbeddedFontHeader.UpgradeToVersion22(FontData)).IsNull();
        await Assert.That(EmbeddedFontHeader.UpgradeToVersion22(new byte[200])).IsNull();
    }

    internal static byte[] CreateEot(uint version, byte[] fontData, byte[]? afterFullName = null)
    {
        using var stream = new MemoryStream();
        using var writer = new BinaryWriter(stream, Encoding.Unicode);
        writer.Write(0u);
        writer.Write((uint)fontData.Length);
        writer.Write(version);
        writer.Write(0u);
        writer.Write(new byte[10]);
        writer.Write((byte)1);
        writer.Write((byte)0);
        writer.Write(400u);
        writer.Write((ushort)0);
        writer.Write((ushort)0x504C);
        writer.Write(new byte[16 + 8 + 4 + 16]);
        string[] names = ["Test Family", "Regular", "Version 1.0", "Test Family Regular"];
        foreach (var name in names)
        {
            writer.Write((ushort)0);
            var bytes = Encoding.Unicode.GetBytes(name);
            writer.Write((ushort)bytes.Length);
            writer.Write(bytes);
        }
        writer.Write(afterFullName ?? []);
        writer.Write(fontData);
        writer.Flush();

        var eot = stream.ToArray();
        BinaryPrimitives.WriteUInt32LittleEndian(eot, (uint)eot.Length);
        return eot;
    }
}
