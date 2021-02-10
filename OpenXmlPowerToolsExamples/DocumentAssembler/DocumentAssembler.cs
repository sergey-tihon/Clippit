﻿// Copyright (c) Microsoft. All rights reserved.
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

using System;
using System.IO;
using System.Xml.Linq;
using Clippit.Word;


namespace Clippit
{
    class Program
    {
        static void Main(string[] args)
        {
            if (args.Length != 3)
            {
                PrintUsage();
                Environment.Exit(0);
            }

            FileInfo templateDoc = new FileInfo(args[0]);
            if (!templateDoc.Exists)
            {
                Console.WriteLine("Error, {0} does not exist.", args[0]);
                PrintUsage();
                Environment.Exit(0);
            }
            FileInfo dataFile = new FileInfo(args[1]);
            if (!dataFile.Exists)
            {
                Console.WriteLine("Error, {0} does not exist.", args[1]);
                PrintUsage();
                Environment.Exit(0);
            }
            FileInfo assembledDoc = new FileInfo(args[2]);
            if (assembledDoc.Exists)
            {
                Console.WriteLine("Error, {0} exists.", args[2]);
                PrintUsage();
                Environment.Exit(0);
            }

            WmlDocument wmlDoc = new WmlDocument(templateDoc.FullName);
            XElement data = XElement.Load(dataFile.FullName);
            bool templateError;
            WmlDocument wmlAssembledDoc = DocumentAssembler.AssembleDocument(wmlDoc, data, out templateError);
            if (templateError)
            {
                Console.WriteLine("Errors in template.");
                Console.WriteLine("See {0} to determine the errors in the template.", assembledDoc.Name);
            }

            wmlAssembledDoc.SaveAs(assembledDoc.FullName);
        }

        static void PrintUsage()
        {
            Console.WriteLine("Usage: DocumentAssembler TemplateDocument.docx Data.xml AssembledDoc.docx");
        }
    }
}
