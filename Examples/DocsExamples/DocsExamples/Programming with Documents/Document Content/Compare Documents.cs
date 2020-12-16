﻿using System;
using Aspose.Words;
using NUnit.Framework;

namespace DocsExamples.Programming_with_Documents.Document_Content
{
    internal class CompareDocument : DocsExamplesBase
    {
        [Test]
        public static void CompareForEqual()
        {
            //ExStart:CompareForEqual
            Document docA = new Document(MyDir + "Document.docx");
            Document docB = docA.Clone();
            
            // DocA now contains changes as revisions.
            docA.Compare(docB, "user", DateTime.Now);

            Console.WriteLine(docA.Revisions.Count == 0 ? "Documents are equal" : "Documents are not equal");
            //ExEnd:CompareForEqual                     
        }

        [Test]
        public static void CompareOptions()
        {
            //ExStart:CompareOptions
            Document docA = new Document(MyDir + "Document.docx");
            Document docB = docA.Clone();

            CompareOptions options = new CompareOptions
            {
                IgnoreFormatting = true,
                IgnoreHeadersAndFooters = true,
                IgnoreCaseChanges = true,
                IgnoreTables = true,
                IgnoreFields = true,
                IgnoreComments = true,
                IgnoreTextboxes = true,
                IgnoreFootnotes = true
            };

            docA.Compare(docB, "user", DateTime.Now, options);

            Console.WriteLine(docA.Revisions.Count == 0 ? "Documents are equal" : "Documents are not equal");
            //ExEnd:CompareOptions                     
        }

        [Test]
        public static void ComparisonTarget()
        {
            //ExStart:ComparisonTarget
            Document docA = new Document(MyDir + "Document.docx");
            Document docB = docA.Clone();

            // Relates to Microsoft Word "Show changes in" option in "Compare Documents" dialog box.
            CompareOptions options = new CompareOptions { IgnoreFormatting = true, Target = ComparisonTargetType.New };

            docA.Compare(docB, "user", DateTime.Now, options);
            //ExEnd:ComparisonTarget
        }

        [Test]
        public static void ComparisonGranularity()
        {
            // ExStart:ComparisonGranularity
            DocumentBuilder builderA = new DocumentBuilder(new Document());
            DocumentBuilder builderB = new DocumentBuilder(new Document());

            builderA.Writeln("This is A simple word");
            builderB.Writeln("This is B simple words");

            CompareOptions compareOptions = new CompareOptions { Granularity = Granularity.CharLevel };

            builderA.Document.Compare(builderB.Document, "author", DateTime.Now, compareOptions);
            // ExEnd:ComparisonGranularity      
        }
    }
}