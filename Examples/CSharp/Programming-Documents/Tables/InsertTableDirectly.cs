using System.Drawing;
using Aspose.Words.Tables;
using NUnit.Framework;

namespace Aspose.Words.Examples.CSharp.Programming_Documents.Working_with_Tables
{
    class InsertTableDirectly : TestDataHelper
    {
        [Test]
        public static void Run()
        {
            //ExStart:InsertTableDirectly
            Document doc = new Document();
            // We start by creating the table object. Note how we must pass the document object
            // To the constructor of each node. This is because every node we create must belong
            // to some document
            Table table = new Table(doc);
            // Add the table to the document
            doc.FirstSection.Body.AppendChild(table);

            // Here we could call EnsureMinimum to create the rows and cells for us. This method is used
            // To ensure that the specified node is valid, in this case a valid table should have at least one
            // Row and one cell, therefore this method creates them for us

            // Instead we will handle creating the row and table ourselves. This would be the best way to do this
            // If we were creating a table inside an algorithm for example
            Row row = new Row(doc);
            row.RowFormat.AllowBreakAcrossPages = true;
            table.AppendChild(row);

            // We can now apply any auto fit settings
            table.AutoFit(AutoFitBehavior.FixedColumnWidths);

            // Create a cell and add it to the row
            Cell cell = new Cell(doc);
            cell.CellFormat.Shading.BackgroundPatternColor = Color.LightBlue;
            cell.CellFormat.Width = 80;

            // Add a paragraph to the cell as well as a new run with some text
            cell.AppendChild(new Paragraph(doc));
            cell.FirstParagraph.AppendChild(new Run(doc, "Row 1, Cell 1 Text"));

            // Add the cell to the row
            row.AppendChild(cell);

            // We would then repeat the process for the other cells and rows in the table
            // We can also speed things up by cloning existing cells and rows
            row.AppendChild(cell.Clone(false));
            row.LastCell.AppendChild(new Paragraph(doc));
            row.LastCell.FirstParagraph.AppendChild(new Run(doc, "Row 1, Cell 2 Text"));
            
            doc.Save(ArtifactsDir + "InsertTableDirectly.docx");
            //ExEnd:InsertTableDirectly
        }
    }
}