﻿using System;
using Aspose.Words.Replacing;
using NUnit.Framework;

namespace Aspose.Words.Examples.CSharp
{
    internal class FindAndReplace : TestDataHelper
    {
        [Test]
        public static void Run()
        {
            Document doc = new Document(QuickStartDir + "ReplaceSimple.doc");

            // Check the text of the document
            Console.WriteLine("Original document text: " + doc.Range.Text);

            // Replace the text in the document
            doc.Range.Replace("_CustomerName_", "James Bond", new FindReplaceOptions(FindReplaceDirection.Forward));

            // Check the replacement was made
            Console.WriteLine("Document text after replace: " + doc.Range.Text);

            // Save the modified document
            doc.Save(ArtifactsDir + "ReplaceSimple.doc");
        }
    }
}