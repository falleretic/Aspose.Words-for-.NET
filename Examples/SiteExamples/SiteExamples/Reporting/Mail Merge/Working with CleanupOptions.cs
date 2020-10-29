using System.Data;
using System.Diagnostics;
using Aspose.Words;
using Aspose.Words.Fields;
using Aspose.Words.MailMerging;
using NUnit.Framework;

namespace SiteExamples.Reporting.Mail_Merge
{
    class WorkingWithMailMergeCleanupOptions : SiteExamplesBase
    {
        [Test]
        public static void RemoveRowsFromTable()
        {
            //ExStart:RemoveRowsFromTable
            Document doc = new Document(MyDir + "Mail merge destination - Northwind suppliers.docx");
            
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
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            FieldMergeField mergeFieldOption1 = (FieldMergeField) builder.InsertField("MERGEFIELD", "Option_1");
            mergeFieldOption1.FieldName = "Option_1";

            // Here is the complete list of cleanable punctuation marks: ! , . : ; ? ¡ ¿
            builder.Write(" ?  ");

            FieldMergeField mergeFieldOption2 = (FieldMergeField) builder.InsertField("MERGEFIELD", "Option_2");
            mergeFieldOption2.FieldName = "Option_2";

            doc.MailMerge.CleanupOptions = MailMergeCleanupOptions.RemoveEmptyParagraphs;
            // The default value of the option is true which means that the behaviour was changed to mimic MS Word
            // If you rely on the old behavior are able to revert it by setting the option to false
            doc.MailMerge.CleanupParagraphsWithPunctuationMarks = true;

            doc.MailMerge.Execute(new[] { "Option_1", "Option_2" }, new object[] { null, null });

            doc.Save(ArtifactsDir + "MailMerge.RemoveColonBetweenEmptyMergeFields.docx");
            //ExEnd:CleanupParagraphsWithPunctuationMarks
        }

        [Test]
        public static void RemoveUnmergedRegions()
        {
            //ExStart:RemoveUnmergedRegions
            Document doc = new Document(MyDir + "Mail merge destination - Northwind suppliers.docx");

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