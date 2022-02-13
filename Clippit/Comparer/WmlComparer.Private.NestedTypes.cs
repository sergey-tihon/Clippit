// Copyright (c) Microsoft. All rights reserved.
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

using System.Drawing;
using System.Xml.Linq;

namespace Clippit
{
    public static partial class WmlComparer
    {
        private class Atgbw
        {
            public int? Key { get; set; }
            public ComparisonUnitAtom ComparisonUnitAtomMember { get; set; }
            public int NextIndex { get; set; }
        }

        private class ConsolidationInfo
        {
            public string Revisor { get; set; }
            public Color Color { get; set; }
            public XElement RevisionElement { get; set; }
            public bool InsertBefore { get; set; }
            public string RevisionHash { get; set; }
            public XElement[] Footnotes { get; set; }
            public XElement[] Endnotes { get; set; }
            public string RevisionString { get; set; }// for debugging purposes only
        }

        private class RecursionInfo
        {
            public XName ElementName { get; set; }
            public XName[] ChildElementPropertyNames { get; set; }
        }
    }
}
