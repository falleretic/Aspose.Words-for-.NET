using System.Data;
using System.Diagnostics;
using Aspose.Words.MailMerging;
using NUnit.Framework;

namespace Aspose.Words.Examples.CSharp
{
    class RemoveEmptyRegions : TestDataHelper
    {
        [Test]
        public static void Run()
        {
            //ExStart:RemoveUnmergedRegions
            Document doc = new Document(MailMergeDir + "TestFile Empty.doc");

            // Create a dummy data source containing no data
            DataSet data = new DataSet();
            //ExStart:MailMergeCleanupOptions
            // Set the appropriate mail merge clean up options to remove any unused regions from the document
            doc.MailMerge.CleanupOptions = MailMergeCleanupOptions.RemoveUnusedRegions;
            // Doc.MailMerge.CleanupOptions = MailMergeCleanupOptions.RemoveContainingFields;
            // Doc.MailMerge.CleanupOptions |= MailMergeCleanupOptions.RemoveStaticFields;
            // Doc.MailMerge.CleanupOptions |= MailMergeCleanupOptions.RemoveEmptyParagraphs;           
            // Doc.MailMerge.CleanupOptions |= MailMergeCleanupOptions.RemoveUnusedFields;
            //ExEnd:MailMergeCleanupOptions
            // Execute mail merge which will have no effect as there is no data. However the regions found in the document will be removed
            // Automatically as they are unused
            doc.MailMerge.ExecuteWithRegions(data);

            doc.Save(ArtifactsDir + "RemoveEmptyRegions.docx");
            //ExEnd:RemoveUnmergedRegions
            Debug.Assert(doc.MailMerge.GetFieldNames().Length == 0,
                "Error: There are still unused regions remaining in the document");
        }
    }
}