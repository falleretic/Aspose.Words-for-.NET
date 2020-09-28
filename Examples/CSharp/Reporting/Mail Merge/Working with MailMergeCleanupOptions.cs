using System.Data;
using System.Diagnostics;
using Aspose.Words.MailMerging;
using NUnit.Framework;

namespace Aspose.Words.Examples.CSharp
{
    class WorkingWithMailMergeCleanupOptions : TestDataHelper
    {
        [Test]
        public static void RemoveRowsFromTable()
        {
            //ExStart:RemoveRowsFromTable
            Document doc = new Document(MailMergeDir + "Mail merge destinations - Northwind traders.docx");
            
            DataSet data = new DataSet();
            doc.MailMerge.CleanupOptions = MailMergeCleanupOptions.RemoveUnusedRegions |
                                           MailMergeCleanupOptions.RemoveEmptyTableRows;
            doc.MailMerge.MergeDuplicateRegions = true;
            doc.MailMerge.ExecuteWithRegions(data);

            doc.Save(ArtifactsDir + "RemoveRowsFromTable.docx");
            //ExEnd:RemoveRowsFromTable
        }

        [Test]
        public static void CleanupParagraphsWithPunctuationMarks()
        {
            //ExStart:CleanupParagraphsWithPunctuationMarks
            Document doc = new Document(MailMergeDir + "Mail merge destinations - Cleanup punctuation marks.docx");

            doc.MailMerge.CleanupOptions = MailMergeCleanupOptions.RemoveEmptyParagraphs;
            doc.MailMerge.CleanupParagraphsWithPunctuationMarks = false;

            doc.MailMerge.Execute(new string[] { "field1", "field2" }, new object[] { "", "" });

            doc.Save(ArtifactsDir + "MailMerge.CleanupPunctuationMarks.docx");
            //ExEnd:CleanupParagraphsWithPunctuationMarks
        }

        [Test]
        public static void Run()
        {
            //ExStart:RemoveUnmergedRegions
            Document doc = new Document(MailMergeDir + "Mail merge destinations - Northwind traders.docx");

            // Create a dummy data source containing no data
            DataSet data = new DataSet();
            //ExStart:MailMergeCleanupOptions
            // Set the appropriate mail merge clean up options to remove any unused regions from the document
            doc.MailMerge.CleanupOptions = MailMergeCleanupOptions.RemoveUnusedRegions;
            // doc.MailMerge.CleanupOptions = MailMergeCleanupOptions.RemoveContainingFields;
            // doc.MailMerge.CleanupOptions |= MailMergeCleanupOptions.RemoveStaticFields;
            // doc.MailMerge.CleanupOptions |= MailMergeCleanupOptions.RemoveEmptyParagraphs;           
            // doc.MailMerge.CleanupOptions |= MailMergeCleanupOptions.RemoveUnusedFields;
            //ExEnd:MailMergeCleanupOptions
            // Execute mail merge which will have no effect as there is no data. However the regions found in the document will be removed
            // Automatically as they are unused
            doc.MailMerge.ExecuteWithRegions(data);

            doc.Save(ArtifactsDir + "MailMerge.RemoveEmptyRegions.docx");
            //ExEnd:RemoveUnmergedRegions
            Debug.Assert(doc.MailMerge.GetFieldNames().Length == 0,
                "Error: There are still unused regions remaining in the document");
        }
    }
}