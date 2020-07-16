using System.Diagnostics;
using Aspose.Words.Tables;
using NUnit.Framework;

namespace Aspose.Words.Examples.CSharp.Tables
{
    class AutoFitTableToContents : TestDataHelper
    {
        [Test]
        public static void Run()
        {
            //ExStart:AutoFitTableToContents
            Document doc = new Document(TablesDir + "TestFile.doc");

            Table table = (Table) doc.GetChild(NodeType.Table, 0, true);
            // Auto fit the table to the cell contents
            table.AutoFit(AutoFitBehavior.AutoFitToContents);

            doc.Save(ArtifactsDir + "AutoFitTableToContents.docx");

            Debug.Assert(doc.FirstSection.Body.Tables[0].PreferredWidth.Type == PreferredWidthType.Auto,
                "PreferredWidth type is not auto");
            Debug.Assert(
                doc.FirstSection.Body.Tables[0].FirstRow.FirstCell.CellFormat.PreferredWidth.Type ==
                PreferredWidthType.Auto, "PrefferedWidth on cell is not auto");
            Debug.Assert(doc.FirstSection.Body.Tables[0].FirstRow.FirstCell.CellFormat.PreferredWidth.Value == 0,
                "PreferredWidth value is not 0");
            //ExEnd:AutoFitTableToContents
        }
    }
}