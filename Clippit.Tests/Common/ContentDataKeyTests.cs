// Copyright (c) Microsoft. All rights reserved.
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

using System.Collections.Generic;
using Clippit;

namespace Clippit.Tests.Common;

/// <summary>
/// Unit tests for <see cref="ContentDataKey"/> — the deduplication key struct
/// used by <see cref="ImageData"/> and <see cref="MediaData"/> in the
/// PresentationBuilder media cache.
/// </summary>
public class ContentDataKeyTests
{
    private static byte[] SomeHash(byte seed = 0) => Enumerable.Range(0, 32).Select(i => (byte)(i + seed)).ToArray();

    // ── CDK001: same content type and same hash → equal ─────────────────────

    [Test]
    public async Task CDK001_SameContentTypeAndHash_Equal()
    {
        var hash = SomeHash();
        var a = new ContentDataKey("image/png", hash);
        var b = new ContentDataKey("image/png", (byte[])hash.Clone());

        await Assert.That(a.Equals(b)).IsTrue();
    }

    // ── CDK002: different content types → not equal ──────────────────────────

    [Test]
    public async Task CDK002_DifferentContentType_NotEqual()
    {
        var hash = SomeHash();
        var a = new ContentDataKey("image/png", hash);
        var b = new ContentDataKey("image/jpeg", hash);

        await Assert.That(a.Equals(b)).IsFalse();
    }

    // ── CDK003: different hash bytes → not equal ─────────────────────────────

    [Test]
    public async Task CDK003_DifferentHash_NotEqual()
    {
        var a = new ContentDataKey("image/png", SomeHash(seed: 0));
        var b = new ContentDataKey("image/png", SomeHash(seed: 1));

        await Assert.That(a.Equals(b)).IsFalse();
    }

    // ── CDK004: Equals(object) overload works correctly ─────────────────────

    [Test]
    public async Task CDK004_EqualsObject_BoxedCopy_IsEqual()
    {
        var hash = SomeHash();
        var a = new ContentDataKey("video/mp4", hash);
        object boxed = new ContentDataKey("video/mp4", (byte[])hash.Clone());

        await Assert.That(a.Equals(boxed)).IsTrue();
    }

    [Test]
    public async Task CDK005_EqualsObject_WrongType_IsFalse()
    {
        var a = new ContentDataKey("image/png", SomeHash());

        await Assert.That(a.Equals("not a key")).IsFalse();
        await Assert.That(a.Equals(null)).IsFalse();
    }

    // ── CDK006: GetHashCode is deterministic ────────────────────────────────

    [Test]
    public async Task CDK006_GetHashCode_SameInputs_SameCode()
    {
        var hash = SomeHash();
        var a = new ContentDataKey("image/png", hash);
        var b = new ContentDataKey("image/png", (byte[])hash.Clone());

        await Assert.That(a.GetHashCode()).IsEqualTo(b.GetHashCode());
    }

    // ── CDK007: different content types produce different hash codes (usually)

    [Test]
    public async Task CDK007_GetHashCode_DifferentContentType_DifferentCode()
    {
        var hash = SomeHash();
        var a = new ContentDataKey("image/png", hash);
        var b = new ContentDataKey("image/jpeg", hash);

        // Hash collisions are theoretically possible but extremely unlikely with SHA-256 seeds.
        await Assert.That(a.GetHashCode()).IsNotEqualTo(b.GetHashCode());
    }

    // ── CDK008: usable as a Dictionary key ──────────────────────────────────

    [Test]
    public async Task CDK008_UsedAsDictionaryKey_LooksUpCorrectly()
    {
        var hash = SomeHash();
        var key1 = new ContentDataKey("image/png", hash);
        var key2 = new ContentDataKey("image/png", (byte[])hash.Clone()); // same value, different array

        var dict = new Dictionary<ContentDataKey, string> { [key1] = "image-part" };

        await Assert.That(dict.ContainsKey(key2)).IsTrue();
        await Assert.That(dict[key2]).IsEqualTo("image-part");
    }

    // ── CDK009: symmetry of Equals ─────────────────────────────────────────

    [Test]
    public async Task CDK009_Equals_IsSymmetric()
    {
        var a = new ContentDataKey("image/png", SomeHash(0));
        var b = new ContentDataKey("image/png", SomeHash(0));

        await Assert.That(a.Equals(b)).IsTrue();
        await Assert.That(b.Equals(a)).IsTrue();
    }
}
