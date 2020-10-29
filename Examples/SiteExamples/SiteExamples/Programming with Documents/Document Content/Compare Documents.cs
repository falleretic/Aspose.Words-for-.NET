using System;
using Aspose.Words;
using NUnit.Framework;

namespace SiteExamples.Programming_with_Documents.Document_Content
{
    class CompareDocument : SiteExamplesBase
    {
        [Test]
        public static void CompareForEqual()
        {
            //ExStart:CompareForEqual
            Document docA = new Document(MyDir + "Document.docx");
            Document docB = docA.Clone();
            
            // DocA now contains changes as revisions
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

            CompareOptions options = new CompareOptions();
            options.IgnoreFormatting = true;
            options.IgnoreHeadersAndFooters = true;
            options.IgnoreCaseChanges = true;
            options.IgnoreTables = true;
            options.IgnoreFields = true;
            options.IgnoreComments = true;
            options.IgnoreTextboxes = true;
            options.IgnoreFootnotes = true;

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

            CompareOptions options = new CompareOptions();
            options.IgnoreFormatting = true;
            // Relates to Microsoft Word "Show changes in" option in "Compare Documents" dialog box.
            options.Target = ComparisonTargetType.New;

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

            CompareOptions compareOptions = new CompareOptions();
            compareOptions.Granularity = Granularity.CharLevel;

            builderA.Document.Compare(builderB.Document, "author", DateTime.Now, compareOptions);
            // ExEnd:ComparisonGranularity      
        }
    }
}