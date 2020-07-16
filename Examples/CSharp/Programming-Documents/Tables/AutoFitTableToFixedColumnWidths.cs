using System.Diagnostics;
using Aspose.Words.Tables;
using NUnit.Framework;

namespace Aspose.Words.Examples.CSharp.Tables
{
    class AutoFitTableToFixedColumnWidths : TestDataHelper
    {
        [Test]
        public static void Run()
        {
            //ExStart:AutoFitTableToFixedColumnWidths
            Document doc = new Document(TablesDir + "TestFile.doc");

            Table table = (Table) doc.GetChild(NodeType.Table, 0, true);
            // Disable autofitting on this table
            table.AutoFit(AutoFitBehavior.FixedColumnWidths);

            doc.Save(ArtifactsDir + "AutoFitTableToFixedColumnWidths.docx");
            
            Debug.Assert(doc.FirstSection.Body.Tables[0].PreferredWidth.Type == PreferredWidthType.Auto,
                "PreferredWidth type is not auto");
            Debug.Assert(doc.FirstSection.Body.Tables[0].PreferredWidth.Value == 0, "PreferredWidth value is not 0");
            Debug.Assert(doc.FirstSection.Body.Tables[0].FirstRow.FirstCell.CellFormat.Width == 69.2,
                "Cell width is not correct.");
            //ExEnd:AutoFitTableToFixedColumnWidths
        }
    }
}