// Copyright (c) Microsoft. All rights reserved.
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

using System;
using System.Collections.Generic;
using System.IO;
using System.IO.Packaging;
using System.Linq;
using System.Xml.Linq;
using Clippit.Internal;

namespace Clippit
{
    public static partial class WmlComparer
    {
        private static XElement MoveRelatedPartsToDestination(
            PackagePart partOfDeletedContent,
            PackagePart partInNewDocument,
            XElement contentElement)
        {
            var elementsToUpdate = contentElement
                .Descendants()
                .Where(d => d.Attributes().Any(a => ComparisonUnitWord.RelationshipAttributeNames.Contains(a.Name)))
                .ToList();

            foreach (var element in elementsToUpdate)
            {
                var attributesToUpdate = element
                    .Attributes()
                    .Where(a => ComparisonUnitWord.RelationshipAttributeNames.Contains(a.Name))
                    .ToList();

                foreach (var att in attributesToUpdate)
                {
                    var rId = (string) att;

                    var relationshipForDeletedPart = partOfDeletedContent.GetRelationship(rId);

                    var targetUri = PackUriHelper
                        .ResolvePartUri(
                            new Uri(partOfDeletedContent.Uri.ToString(), UriKind.Relative),
                            relationshipForDeletedPart.TargetUri);

                    var relatedPackagePart = partOfDeletedContent.Package.GetPart(targetUri);
                    var uriSplit = relatedPackagePart.Uri.ToString().Split('/');
                    var last = uriSplit[^1].Split('.');
                    string uriString;
                    if (last.Length == 2)
                    {
                        uriString = uriSplit.SkipLast(1).Select(p => p + "/").StringConcatenate() +
                                    "P" + Guid.NewGuid().ToString().Replace("-", "") + "." + last[1];
                    }
                    else
                    {
                        uriString = uriSplit.SkipLast(1).Select(p => p + "/").StringConcatenate() +
                                    "P" + Guid.NewGuid().ToString().Replace("-", "");
                    }

                    var uri = relatedPackagePart.Uri.IsAbsoluteUri
                        ? new Uri(uriString, UriKind.Absolute)
                        : new Uri(uriString, UriKind.Relative);

                    var newPart = partInNewDocument.Package.CreatePart(uri, relatedPackagePart.ContentType);

                    // ReSharper disable once PossibleNullReferenceException
                    using (var oldPartStream = relatedPackagePart.GetStream())
                    using (var newPartStream = newPart.GetStream())
                    {
                        oldPartStream.CopyTo(newPartStream);
                    }

                    var newRid = Relationships.GetNewRelationshipId();
                    partInNewDocument.CreateRelationship(newPart.Uri, TargetMode.Internal,
                        relationshipForDeletedPart.RelationshipType, newRid);
                    att.Value = newRid;

                    if (newPart.ContentType.EndsWith("xml"))
                    {
                        XDocument newPartXDoc;
                        using (var stream = newPart.GetStream())
                        {
                            newPartXDoc = XDocument.Load(stream);
                            MoveRelatedPartsToDestination(relatedPackagePart, newPart, newPartXDoc.Root);
                        }

                        using (var stream = newPart.GetStream())
                            newPartXDoc.Save(stream);
                    }
                }
            }

            return contentElement;
        }

        private static XAttribute GetXmlSpaceAttribute(string textOfTextElement)
        {
            if (char.IsWhiteSpace(textOfTextElement[0]) ||
                char.IsWhiteSpace(textOfTextElement[^1]))
                return new XAttribute(XNamespace.Xml + "space", "preserve");

            return null;
        }
    }
}
