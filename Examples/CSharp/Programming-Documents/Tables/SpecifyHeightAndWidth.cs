using System.Drawing;
using Aspose.Words.Tables;
using NUnit.Framework;

namespace Aspose.Words.Examples.CSharp.Programming_Documents.Working_with_Tables
{
    class SpecifyHeightAndWidth : TestDataHelper
    {
        /// <summary>
        /// Shows how to set a table to auto fit to 50% of the page width.
        /// </summary>
        [Test]
        public static void AutoFitToPageWidth()
        {
            //ExStart:AutoFitToPageWidth
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Insert a table with a width that takes up half the page width
            Table table = builder.StartTable();

            // Insert a few cells
            builder.InsertCell();
            table.PreferredWidth = PreferredWidth.FromPercent(50);
            builder.Writeln("Cell #1");

            builder.InsertCell();
            builder.Writeln("Cell #2");

            builder.InsertCell();
            builder.Writeln("Cell #3");

            doc.Save(ArtifactsDir + "AutoFitToPageWidth.docx");
            //ExEnd:AutoFitToPageWidth
        }

        /// <summary>
        /// Shows how to set the different preferred width settings.
        /// </summary>
        [Test]
        public static void SetPreferredWidthSettings()
        {
            //ExStart:SetPreferredWidthSettings
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Insert a table row made up of three cells which have different preferred widths
            builder.StartTable();

            // Insert an absolute sized cell
            builder.InsertCell();
            builder.CellFormat.PreferredWidth = PreferredWidth.FromPoints(40);
            builder.CellFormat.Shading.BackgroundPatternColor = Color.LightYellow;
            builder.Writeln("Cell at 40 points width");

            // Insert a relative (percent) sized cell
            builder.InsertCell();
            builder.CellFormat.PreferredWidth = PreferredWidth.FromPercent(20);
            builder.CellFormat.Shading.BackgroundPatternColor = Color.LightBlue;
            builder.Writeln("Cell at 20% width");

            // Insert a auto sized cell
            builder.InsertCell();
            builder.CellFormat.PreferredWidth = PreferredWidth.Auto;
            builder.CellFormat.Shading.BackgroundPatternColor = Color.LightGreen;
            builder.Writeln(
                "Cell automatically sized. The size of this cell is calculated from the table preferred width.");
            builder.Writeln("In this case the cell will fill up the rest of the available space.");

            doc.Save(ArtifactsDir + "SetPreferredWidthSettings.docx");
            //ExEnd:SetPreferredWidthSettings
        }

        /// <summary>
        /// Shows how to retrieves the preferred width type of a table cell.
        /// </summary>
        [Test]
        public static void RetrievePreferredWidthType()
        {
            //ExStart:RetrievePreferredWidthType
            Document doc = new Document(TablesDir + "Table.SimpleTable.doc");

            // Retrieve the first table in the document
            Table table = (Table) doc.GetChild(NodeType.Table, 0, true);
            //ExStart:AllowAutoFit
            table.AllowAutoFit = true;
            //ExEnd:AllowAutoFit

            Cell firstCell = table.FirstRow.FirstCell;
            PreferredWidthType type = firstCell.CellFormat.PreferredWidth.Type;
            double value = firstCell.CellFormat.PreferredWidth.Value;
            //ExEnd:RetrievePreferredWidthType
        }
    }
}