using System;
using System.Drawing;
using Aspose.Words.Saving;
using Aspose.Words.Tables;
using NUnit.Framework;

namespace Aspose.Words.Examples.CSharp.Tables
{
    class ApplyFormatting : TestDataHelper
    {
        /// <summary>
        /// Shows how to get distance between table surrounding text.
        /// </summary>
        [Test]
        public static void GetDistanceBetweenTableSurroundingText()
        {
            //ExStart:GetDistancebetweenTableSurroundingText
            Document doc = new Document(TablesDir + "Table.EmptyTable.doc");

            Console.WriteLine("\nGet distance between table left, right, bottom, top and the surrounding text.");
            Table table = (Table) doc.GetChild(NodeType.Table, 0, true);

            Console.WriteLine(table.DistanceTop);
            Console.WriteLine(table.DistanceBottom);
            Console.WriteLine(table.DistanceRight);
            Console.WriteLine(table.DistanceLeft);
            //ExEnd:GetDistancebetweenTableSurroundingText
        }

        /// <summary>
        /// Shows how to apply outline border to a table.
        /// </summary>
        [Test]
        public static void ApplyOutlineBorder()
        {
            //ExStart:ApplyOutlineBorder
            Document doc = new Document(TablesDir + "Table.EmptyTable.doc");

            Table table = (Table) doc.GetChild(NodeType.Table, 0, true);
            // Align the table to the center of the page
            table.Alignment = TableAlignment.Center;
            // Clear any existing borders from the table
            table.ClearBorders();

            // Set a green border around the table but not inside
            table.SetBorder(BorderType.Left, LineStyle.Single, 1.5, Color.Green, true);
            table.SetBorder(BorderType.Right, LineStyle.Single, 1.5, Color.Green, true);
            table.SetBorder(BorderType.Top, LineStyle.Single, 1.5, Color.Green, true);
            table.SetBorder(BorderType.Bottom, LineStyle.Single, 1.5, Color.Green, true);

            // Fill the cells with a light green solid color
            table.SetShading(TextureIndex.TextureSolid, Color.LightGreen, Color.Empty);

            doc.Save(ArtifactsDir + "ApplyOutlineBorder.docx");
            //ExEnd:ApplyOutlineBorder
        }

        /// <summary>
        /// Shows how to build a table with all borders enabled (grid).
        /// </summary>
        [Test]
        public static void BuildTableWithBordersEnabled()
        {
            //ExStart:BuildTableWithBordersEnabled
            Document doc = new Document(TablesDir + "Table.EmptyTable.doc");

            Table table = (Table) doc.GetChild(NodeType.Table, 0, true);
            // Clear any existing borders from the table
            table.ClearBorders();
            // Set a green border around and inside the table
            table.SetBorders(LineStyle.Single, 1.5, Color.Green);

            doc.Save(ArtifactsDir + "BuildTableWithBordersEnabled.docx");
            //ExEnd:BuildTableWithBordersEnabled
        }

        /// <summary>
        /// Shows how to modify formatting of a table row.
        /// </summary>
        [Test]
        public static void ModifyRowFormatting()
        {
            //ExStart:ModifyRowFormatting
            Document doc = new Document(TablesDir + "Table.Document.doc");

            Table table = (Table) doc.GetChild(NodeType.Table, 0, true);
            // Retrieve the first row in the table
            Row firstRow = table.FirstRow;
            // Modify some row level properties
            firstRow.RowFormat.Borders.LineStyle = LineStyle.None;
            firstRow.RowFormat.HeightRule = HeightRule.Auto;
            firstRow.RowFormat.AllowBreakAcrossPages = true;
            //ExEnd:ModifyRowFormatting
        }

        /// <summary>
        /// Shows how to create a table that contains a single cell and apply row formatting.
        /// </summary>
        [Test]
        public static void ApplyRowFormatting()
        {
            //ExStart:ApplyRowFormatting
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            Table table = builder.StartTable();
            builder.InsertCell();

            // Set the row formatting
            RowFormat rowFormat = builder.RowFormat;
            rowFormat.Height = 100;
            rowFormat.HeightRule = HeightRule.Exactly;
            // These formatting properties are set on the table and are applied to all rows in the table
            table.LeftPadding = 30;
            table.RightPadding = 30;
            table.TopPadding = 30;
            table.BottomPadding = 30;

            builder.Writeln("I'm a wonderful formatted row.");

            builder.EndRow();
            builder.EndTable();

            doc.Save(ArtifactsDir + "ApplyRowFormatting.docx");
            //ExEnd:ApplyRowFormatting
        }

        /// <summary>
        /// Shows how to modify formatting of a table cell.
        /// </summary>
        [Test]
        public static void SetCellPadding()
        {
            //ExStart:SetCellPadding
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            builder.StartTable();
            builder.InsertCell();

            // Sets the amount of space (in points) to add to the left/top/right/bottom of the contents of cell
            builder.CellFormat.SetPaddings(30, 50, 30, 50);
            builder.Writeln("I'm a wonderful formatted cell.");

            builder.EndRow();
            builder.EndTable();

            doc.Save(ArtifactsDir + "SetCellPadding.docx");
            //ExEnd:SetCellPadding
        }

        /// <summary>
        /// Shows how to modify formatting of a table cell.
        /// </summary>
        [Test]
        public static void ModifyCellFormatting()
        {
            //ExStart:ModifyCellFormatting
            Document doc = new Document(TablesDir + "Table.Document.doc");
            Table table = (Table) doc.GetChild(NodeType.Table, 0, true);

            // Retrieve the first cell in the table
            Cell firstCell = table.FirstRow.FirstCell;
            // Modify some cell level properties
            firstCell.CellFormat.Width = 30; // In points
            firstCell.CellFormat.Orientation = TextOrientation.Downward;
            firstCell.CellFormat.Shading.ForegroundPatternColor = Color.LightGreen;
            //ExEnd:ModifyCellFormatting
        }

        /// <summary>
        /// Shows how to format table and cell with different borders and shadings.
        /// </summary>
        [Test]
        public static void FormatTableAndCellWithDifferentBorders()
        {
            //ExStart:FormatTableAndCellWithDifferentBorders
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            Table table = builder.StartTable();
            builder.InsertCell();

            // Set the borders for the entire table
            table.SetBorders(LineStyle.Single, 2.0, Color.Black);
            // Set the cell shading for this cell.
            builder.CellFormat.Shading.BackgroundPatternColor = Color.Red;
            builder.Writeln("Cell #1");

            builder.InsertCell();
            // Specify a different cell shading for the second cell
            builder.CellFormat.Shading.BackgroundPatternColor = Color.Green;
            builder.Writeln("Cell #2");

            // End this row
            builder.EndRow();

            // Clear the cell formatting from previous operations
            builder.CellFormat.ClearFormatting();

            // Create the second row
            builder.InsertCell();

            // Create larger borders for the first cell of this row. This will be different
            // Compared to the borders set for the table
            builder.CellFormat.Borders.Left.LineWidth = 4.0;
            builder.CellFormat.Borders.Right.LineWidth = 4.0;
            builder.CellFormat.Borders.Top.LineWidth = 4.0;
            builder.CellFormat.Borders.Bottom.LineWidth = 4.0;
            builder.Writeln("Cell #3");

            builder.InsertCell();
            // Clear the cell formatting from the previous cell
            builder.CellFormat.ClearFormatting();
            builder.Writeln("Cell #4");
            
            doc.Save(ArtifactsDir + "FormatTableAndCellWithDifferentBorders.docx");
            //ExEnd:FormatTableAndCellWithDifferentBorders
        }

        /// <summary>
        /// Shows how to set title and description of table.
        /// </summary>
        [Test]
        public static void SetTableTitleAndDescription()
        {
            //ExStart:SetTableTitleandDescription
            Document doc = new Document(TablesDir + "Table.Document.doc");
            Table table = (Table) doc.GetChild(NodeType.Table, 0, true);
            table.Title = "Test title";
            table.Description = "Test description";

            OoxmlSaveOptions options = new OoxmlSaveOptions();
            options.Compliance = OoxmlCompliance.Iso29500_2008_Strict;

            doc.CompatibilityOptions.OptimizeFor(Settings.MsWordVersion.Word2016);

            doc.Save(ArtifactsDir + "SetTableTitleAndDescription.docx", options);
            //ExEnd:SetTableTitleandDescription
        }

        /// <summary>
        /// Shows how to set "Allow spacing between cells" option
        /// </summary>
        [Test]
        public static void AllowCellSpacing()
        {
            //ExStart:AllowCellSpacing
            Document doc = new Document(TablesDir + "Table.Document.doc");

            Table table = (Table) doc.GetChild(NodeType.Table, 0, true);
            table.AllowCellSpacing = true;
            table.CellSpacing = 2;
            
            doc.Save(ArtifactsDir + "AllowCellSpacing.docx");
            //ExEnd:AllowCellSpacing
        }
    }
}