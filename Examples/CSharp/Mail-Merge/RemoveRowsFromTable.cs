using System.Data;
using Aspose.Words.MailMerging;
using NUnit.Framework;

namespace Aspose.Words.Examples.CSharp
{
    class RemoveRowsFromTable : TestDataHelper
    {
        [Test]
        public static void Run()
        {
            //ExStart:RemoveRowsFromTable
            Document doc = new Document(MailMergeDir + "RemoveTableRows.doc");
            DataSet data = new DataSet();
            doc.MailMerge.CleanupOptions = MailMergeCleanupOptions.RemoveUnusedRegions |
                                           MailMergeCleanupOptions.RemoveEmptyTableRows;
            doc.MailMerge.MergeDuplicateRegions = true;
            doc.MailMerge.ExecuteWithRegions(data);

            doc.Save(ArtifactsDir + "RemoveRowsFromTable.docx");
            //ExEnd:RemoveRowsFromTable
        }
    }
}