using System;

namespace Aspose.Words.Examples.CSharp.Programming_Documents.Working_With_Document
{
    class CompareDocument : TestDataHelper
    {
        public static void Run()
        {
            NormalComparison();
            CompareForEqual();
            CompareDocumentWithCompareOptions();
            CompareDocumentWithComparisonTarget();
        }

        private static void NormalComparison()
        {
            //ExStart:NormalComparison
            Document docA = new Document(DocumentDir + "TestFile.doc");
            Document docB = new Document(DocumentDir + "TestFile - Copy.doc");
            
            // DocA now contains changes as revisions
            docA.Compare(docB, "user", DateTime.Now);
            //ExEnd:NormalComparison                     
        }

        private static void CompareForEqual()
        {
            //ExStart:CompareForEqual
            Document docA = new Document(DocumentDir + "TestFile.doc");
            Document docB = new Document(DocumentDir + "TestFile - Copy.doc");
            
            // DocA now contains changes as revisions
            docA.Compare(docB, "user", DateTime.Now);
            Console.WriteLine(docA.Revisions.Count == 0 ? "Documents are equal" : "Documents are not equal");
            //ExEnd:CompareForEqual                     
        }

        private static void CompareDocumentWithCompareOptions()
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

        private static void CompareDocumentWithComparisonTarget()
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
    }
}