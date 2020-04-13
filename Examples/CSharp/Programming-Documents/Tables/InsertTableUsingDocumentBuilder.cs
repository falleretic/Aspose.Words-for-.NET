using System;
using System.Drawing;
using Aspose.Words.Tables;

namespace Aspose.Words.Examples.CSharp.Programming_Documents.Working_with_Tables
{
    class InsertTableUsingDocumentBuilder : TestDataHelper
    {
        public static void Run()
        {
            SimpleTable();
            FormattedTable();
            NestedTable();
        }

        private static void SimpleTable()
        {
            //ExStart:SimpleTable
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);
            // We call this method to start building the table
            builder.StartTable();
            builder.InsertCell();
            builder.Write("Row 1, Cell 1 Content.");
            // Build the second cell
            builder.InsertCell();
            builder.Write("Row 1, Cell 2 Content.");
            // Call the following method to end the row and start a new row
            builder.EndRow();

            // Build the first cell of the second row
            builder.InsertCell();
            builder.Write("Row 2, Cell 1 Content");

            // Build the second cell
            builder.InsertCell();
            builder.Write("Row 2, Cell 2 Content.");
            builder.EndRow();

            // Signal that we have finished building the table
            builder.EndTable();

            doc.Save(ArtifactsDir + "SimpleTable.docx");
            //ExEnd:SimpleTable
        }

        private static void FormattedTable()
        {
            //ExStart:FormattedTable
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            Table table = builder.StartTable();
            // Make the header row
            builder.InsertCell();

            // Set the left indent for the table. Table wide formatting must be applied after 
            // At least one row is present in the table
            table.LeftIndent = 20.0;

            // Set height and define the height rule for the header row
            builder.RowFormat.Height = 40.0;
            builder.RowFormat.HeightRule = HeightRule.AtLeast;

            // Some special features for the header row
            builder.CellFormat.Shading.BackgroundPatternColor = Color.FromArgb(198, 217, 241);
            builder.ParagraphFormat.Alignment = ParagraphAlignment.Center;
            builder.Font.Size = 16;
            builder.Font.Name = "Arial";
            builder.Font.Bold = true;

            builder.CellFormat.Width = 100.0;
            builder.Write("Header Row,\n Cell 1");

            // We don't need to specify the width of this cell because it's inherited from the previous cell
            builder.InsertCell();
            builder.Write("Header Row,\n Cell 2");

            builder.InsertCell();
            builder.CellFormat.Width = 200.0;
            builder.Write("Header Row,\n Cell 3");
            builder.EndRow();

            // Set features for the other rows and cells
            builder.CellFormat.Shading.BackgroundPatternColor = Color.White;
            builder.CellFormat.Width = 100.0;
            builder.CellFormat.VerticalAlignment = CellVerticalAlignment.Center;

            // Reset height and define a different height rule for table body
            builder.RowFormat.Height = 30.0;
            builder.RowFormat.HeightRule = HeightRule.Auto;
            builder.InsertCell();
            // Reset font formatting
            builder.Font.Size = 12;
            builder.Font.Bold = false;

            // Build the other cells
            builder.Write("Row 1, Cell 1 Content");
            builder.InsertCell();
            builder.Write("Row 1, Cell 2 Content");

            builder.InsertCell();
            builder.CellFormat.Width = 200.0;
            builder.Write("Row 1, Cell 3 Content");
            builder.EndRow();

            builder.InsertCell();
            builder.CellFormat.Width = 100.0;
            builder.Write("Row 2, Cell 1 Content");

            builder.InsertCell();
            builder.Write("Row 2, Cell 2 Content");

            builder.InsertCell();
            builder.CellFormat.Width = 200.0;
            builder.Write("Row 2, Cell 3 Content.");
            builder.EndRow();
            builder.EndTable();

            doc.Save(ArtifactsDir + "FormattedTable.docx");
            //ExEnd:FormattedTable
        }

        private static void NestedTable()
        {
            //ExStart:NestedTable
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Build the outer table
            Cell cell = builder.InsertCell();
            builder.Writeln("Outer Table Cell 1");

            builder.InsertCell();
            builder.Writeln("Outer Table Cell 2");

            // This call is important in order to create a nested table within the first table
            // Without this call the cells inserted below will be appended to the outer table
            builder.EndTable();

            // Move to the first cell of the outer table
            builder.MoveTo(cell.FirstParagraph);

            // Build the inner table
            builder.InsertCell();
            builder.Writeln("Inner Table Cell 1");
            builder.InsertCell();
            builder.Writeln("Inner Table Cell 2");
            builder.EndTable();

            doc.Save(ArtifactsDir + "NestedTable.docx");
            //ExEnd:NestedTable
        }
    }
}