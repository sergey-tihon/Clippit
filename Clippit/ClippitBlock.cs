// Copyright (c) Microsoft. All rights reserved.
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

using System;
using System.Linq;
using System.Xml;
using System.Xml.Linq;
using DocumentFormat.OpenXml.Packaging;

namespace Clippit
{
    /// <summary>
    /// Provides an elegant way of wrapping a set of invocations of the PowerTools in a using
    /// statement that demarcates those invocations as one "block" before and after which the
    /// strongly typed classes provided by the Open XML SDK can be used safely.
    /// </summary>
    /// <remarks>
    /// <para>
    /// This class lends itself to scenarios where the PowerTools and Linq-to-XML are used as
    /// a secondary API for working with Open XML elements, next to the strongly typed classes
    /// provided by the Open XML SDK. In these scenarios, the class would be
    /// used as follows:
    /// </para>
    /// <code>
    ///     [Your code using the strongly typed classes]
    /// 
    ///     using (new PowerToolsBlock(wordprocessingDocument))
    ///     {
    ///         [Your code using the PowerTools]
    ///     }
    /// 
    ///    [Your code using the strongly typed classes]
    /// </code>
    /// <para>
    /// Upon creation, instances of this class will invoke the
    /// <see cref="ClippitBlockExtensions.BeginPowerToolsBlock"/> method on the package
    /// to begin the transaction.  Upon disposal, instances of this class will call the
    /// <see cref="ClippitBlockExtensions.EndPowerToolsBlock"/> method on the package
    /// to end the transaction.
    /// </para>
    /// </remarks>
    /// <seealso cref="StronglyTypedBlock" />
    /// <seealso cref="ClippitBlockExtensions.BeginPowerToolsBlock"/>
    /// <seealso cref="ClippitBlockExtensions.EndPowerToolsBlock"/>
    public class ClippitBlock : IDisposable
    {
        private OpenXmlPackage _package;

        public ClippitBlock(OpenXmlPackage package)
        {
            _package = package ?? throw new ArgumentNullException(nameof(package));
            _package.BeginPowerToolsBlock();
        }

        public void Dispose()
        {
            if (_package is null) return;

            _package.EndPowerToolsBlock();
            _package = null;
        }
    }
    
    public static class ClippitBlockExtensions
    {
        /// <summary>
        /// Begins a PowerTools Block by (1) removing annotations and, unless the package was
        /// opened in read-only mode, (2) saving the package.
        /// </summary>
        /// <remarks>
        /// Removes <see cref="XDocument" /> and <see cref="XmlNamespaceManager" /> instances
        /// added by <see cref="PtOpenXmlExtensions.GetXDocument(OpenXmlPart)" />,
        /// <see cref="PtOpenXmlExtensions.GetXDocument(OpenXmlPart, out XmlNamespaceManager)" />,
        /// <see cref="PtOpenXmlExtensions.PutXDocument(OpenXmlPart)" />,
        /// <see cref="PtOpenXmlExtensions.PutXDocument(OpenXmlPart, XDocument)" />, and
        /// <see cref="PtOpenXmlExtensions.PutXDocumentWithFormatting(OpenXmlPart)" />.
        /// methods.
        /// </remarks>
        /// <param name="package">
        /// A <see cref="WordprocessingDocument" />, <see cref="SpreadsheetDocument" />,
        /// or <see cref="PresentationDocument" />.
        /// </param>
        public static void BeginPowerToolsBlock(this OpenXmlPackage package)
        {
            if (package is null) throw new ArgumentNullException(nameof(package));

            package.RemovePowerToolsAnnotations();
            package.Save();
        }

        /// <summary>
        /// Ends a PowerTools Block by reloading the root elements of all package parts
        /// that were changed by the PowerTools. A part is deemed changed by the PowerTools
        /// if it has an annotation of type <see cref="XDocument" />.
        /// </summary>
        /// <param name="package">
        /// A <see cref="WordprocessingDocument" />, <see cref="SpreadsheetDocument" />,
        /// or <see cref="PresentationDocument" />.
        /// </param>
        public static void EndPowerToolsBlock(this OpenXmlPackage package)
        {
            if (package is null) throw new ArgumentNullException(nameof(package));

            foreach (var part in package.GetAllParts())
            {
                if (part.Annotations<XDocument>().Any() && part.RootElement != null)
                    part.RootElement.Reload();
            }
        }

        private static void RemovePowerToolsAnnotations(this OpenXmlPackage package)
        {
            if (package is null) throw new ArgumentNullException(nameof(package));

            foreach (var part in package.GetAllParts())
            {
                part.RemoveAnnotations<XDocument>();
                part.RemoveAnnotations<XmlNamespaceManager>();
            }
        }
    }
}
