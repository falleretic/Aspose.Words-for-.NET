using System;
using NUnit.Framework;

namespace Aspose.Words.Examples.CSharp.Programming_Documents.Working_With_Document
{
    class CompareDocument : TestDataHelper
    {
        [Test]
        public static void NormalComparison()
        {
            //ExStart:NormalComparison
            Document docA = new Document(DocumentDir + "TestFile.doc");
            Document docB = new Document(DocumentDir + "TestFile - Copy.doc");
            
            // DocA now contains changes as revisions
            docA.Compare(docB, "user", DateTime.Now);
            //ExEnd:NormalComparison                     
        }

        [Test]
        public static void CompareForEqual()
        {
            //ExStart:CompareForEqual
            Document docA = new Document(DocumentDir + "TestFile.doc");
            Document docB = new Document(DocumentDir + "TestFile - Copy.doc");
            
            // DocA now contains changes as revisions
            docA.Compare(docB, "user", DateTime.Now);
            Console.WriteLine(docA.Revisions.Count == 0 ? "Documents are equal" : "Documents are not equal");
            //ExEnd:CompareForEqual                     
        }

        [Test]
        public static void CompareDocumentWithCompareOptions()
        {
            //ExStart:CompareDocumentWithCompareOptions
            Document docA = new Document(DocumentDir + "TestFile.doc");
            Document docB = new Document(DocumentDir + "TestFile - Copy.doc");

            CompareOptions options = new CompareOptions();
            options.IgnoreFormatting = true;
            options.IgnoreHeadersAndFooters = true;
            options.IgnoreCaseChanges = true;
            options.IgnoreTables = true;
            options.IgnoreFields = true;
            options.IgnoreComments = true;
            options.IgnoreTextboxes = true;
            options.IgnoreFootnotes = true;

            // DocA now contains changes as revisions
            docA.Compare(docB, "user", DateTime.Now, options);
            Console.WriteLine(docA.Revisions.Count == 0 ? "Documents are equal" : "Documents are not equal");
            //ExEnd:CompareDocumentWithCompareOptions                     
        }

        [Test]
        public static void CompareDocumentWithComparisonTarget()
        {
            //ExStart:CompareDocumentWithComparisonTarget
            Document docA = new Document(DocumentDir + "TestFile.doc");
            Document docB = new Document(DocumentDir + "TestFile - Copy.doc");

            CompareOptions options = new CompareOptions();
            options.IgnoreFormatting = true;
            // Relates to Microsoft Word "Show changes in" option in "Compare Documents" dialog box
            options.Target = ComparisonTargetType.New;

            docA.Compare(docB, "user", DateTime.Now, options);
            //ExEnd:CompareDocumentWithComparisonTarget
        }

        public static void SpecifyComparisonGranularity(string dataDir)
        {
            // ExStart:SpecifyComparisonGranularity
            DocumentBuilder builderA = new DocumentBuilder(new Document());
            DocumentBuilder builderB = new DocumentBuilder(new Document());

            builderA.Writeln("This is A simple word");
            builderB.Writeln("This is B simple words");

            CompareOptions co = new CompareOptions();
            co.Granularity = Granularity.CharLevel;

            builderA.Document.Compare(builderB.Document, "author", DateTime.Now, co);
            // ExEnd:SpecifyComparisonGranularity      
        }
    }
}