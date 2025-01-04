// Copyright (c) Microsoft. All rights reserved.
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

using System.Xml.Linq;

namespace Clippit
{
    public static partial class WmlComparer
    {
        public class WmlComparerRevision
        {
            public WmlComparerRevisionType RevisionType;
            public string Text;
            public string Author;
            public string Date;
            public XElement ContentXElement;
            public XElement RevisionXElement;
            public Uri PartUri;
            public string PartContentType;
        }

        public enum WmlComparerRevisionType
        {
            Inserted,
            Deleted,
        }
    }
}
