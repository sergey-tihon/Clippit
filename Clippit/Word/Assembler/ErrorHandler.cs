using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml.Linq;

namespace Clippit.Word.Assembler
{
    internal static class ErrorHandler
    {
        internal static object CreateContextErrorMessage(this XElement element, string errorMessage, TemplateError templateError)
        {
            XElement para = element.Descendants(W.p).FirstOrDefault();
            XElement run = element.Descendants(W.r).FirstOrDefault();
            var errorRun = CreateRunErrorMessage(errorMessage, templateError);
            if (para != null)
                return new XElement(W.p, errorRun);
            else
                return errorRun;
        }

        internal static XElement CreateRunErrorMessage(string errorMessage, TemplateError templateError)
        {
            templateError.HasError = true;
            var errorRun = new XElement(W.r,
                new XElement(W.rPr,
                    new XElement(W.color, new XAttribute(W.val, "FF0000")),
                    new XElement(W.highlight, new XAttribute(W.val, "yellow"))),
                    new XElement(W.t, errorMessage));
            return errorRun;
        }

        internal static XElement CreateParaErrorMessage(string errorMessage, TemplateError templateError)
        {
            templateError.HasError = true;
            var errorPara = new XElement(W.p,
                new XElement(W.r,
                    new XElement(W.rPr,
                        new XElement(W.color, new XAttribute(W.val, "FF0000")),
                        new XElement(W.highlight, new XAttribute(W.val, "yellow"))),
                        new XElement(W.t, errorMessage)));
            return errorPara;
        }
    }
}
