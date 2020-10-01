using System;
using System.Collections;
using System.Collections.Generic;
using System.Data;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Text;
using System.Xml;
using Aspose.Words.Drawing;
using Aspose.Words.Replacing;
using Aspose.Words.Tables;
using NUnit.Framework;

namespace Aspose.Words.Examples.CSharp.Programming_with_Documents.Document_Content
{
    class WorkingWithTables : TestDataHelper
    {
        [Test]
        public static void RemoveColumn()
        {
            //ExStart:RemoveColumn
            Document doc = new Document(MyDir + "Tables.docx");

            // Get the second table in the document
            Table table = (Table) doc.GetChild(NodeType.Table, 1, true);

            // Get the third column from the table and remove it
            Column column = Column.FromIndex(table, 2);
            column.Remove();
            //ExEnd:RemoveColumn
        }

        [Test]
        public static void InsertBlankColumn()
        {
            //ExStart:InsertBlankColumn
            Document doc = new Document(MyDir + "Tables.docx");

            // Get the first table in the document
            Table table = (Table) doc.GetChild(NodeType.Table, 0, true);

            //ExStart:GetPlainText
            // Get the second column in the table
            Column column = Column.FromIndex(table, 0);
            // Print the plain text of the column to the screen
            Console.WriteLine(column.ToTxt());
            //ExEnd:GetPlainText
            // Create a new column to the left of this column
            // This is the same as using the "Insert Column Before" command in Microsoft Word
            Column newColumn = column.InsertColumnBefore();

            // Add some text to each of the column cells
            foreach (Cell cell in newColumn.Cells)
                cell.FirstParagraph.AppendChild(new Run(doc, "Column Text " + newColumn.IndexOf(cell)));
            //ExEnd:InsertBlankColumn
        }

        //ExStart:ColumnClass
        /// <summary>
        /// Represents a facade object for a column of a table in a Microsoft Word document.
        /// </summary>
        internal class Column
        {
            private Column(Table table, int columnIndex)
            {
                mTable = table ?? throw new ArgumentException("table");
                mColumnIndex = columnIndex;
            }

            /// <summary>
            /// Returns a new column facade from the table and supplied zero-based index.
            /// </summary>
            public static Column FromIndex(Table table, int columnIndex)
            {
                return new Column(table, columnIndex);
            }

            /// <summary>
            /// Returns the cells which make up the column.
            /// </summary>
            public Cell[] Cells => (Cell[]) GetColumnCells().ToArray(typeof(Cell));

            /// <summary>
            /// Returns the index of the given cell in the column.
            /// </summary>
            public int IndexOf(Cell cell)
            {
                return GetColumnCells().IndexOf(cell);
            }

            /// <summary>
            /// Inserts a brand new column before this column into the table.
            /// </summary>
            public Column InsertColumnBefore()
            {
                Cell[] columnCells = Cells;

                if (columnCells.Length == 0)
                    throw new ArgumentException("Column must not be empty");

                // Create a clone of this column
                foreach (Cell cell in columnCells)
                    cell.ParentRow.InsertBefore(cell.Clone(false), cell);

                // This is the new column
                Column column = new Column(columnCells[0].ParentRow.ParentTable, mColumnIndex);

                // We want to make sure that the cells are all valid to work with (have at least one paragraph)
                foreach (Cell cell in column.Cells)
                    cell.EnsureMinimum();

                // Increase the index which this column represents since there is now one extra column infront
                mColumnIndex++;

                return column;
            }

            /// <summary>
            /// Removes the column from the table.
            /// </summary>
            public void Remove()
            {
                foreach (Cell cell in Cells)
                    cell.Remove();
            }

            /// <summary>
            /// Returns the text of the column. 
            /// </summary>
            public string ToTxt()
            {
                StringBuilder builder = new StringBuilder();

                foreach (Cell cell in Cells)
                    builder.Append(cell.ToString(SaveFormat.Text));

                return builder.ToString();
            }

            /// <summary>
            /// Provides an up-to-date collection of cells which make up the column represented by this facade.
            /// </summary>
            private ArrayList GetColumnCells()
            {
                ArrayList columnCells = new ArrayList();

                foreach (Row row in mTable.Rows)
                {
                    Cell cell = row.Cells[mColumnIndex];
                    if (cell != null)
                        columnCells.Add(cell);
                }

                return columnCells;
            }

            private int mColumnIndex;
            private readonly Table mTable;
        }
        //ExEnd:ColumnClass

        [Test]
        public static void AutoFitTableToContents()
        {
            //ExStart:AutoFitTableToContents
            Document doc = new Document(MyDir + "Tables.docx");

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

        [Test]
        public static void AutoFitTableToFixedColumnWidths()
        {
            //ExStart:AutoFitTableToFixedColumnWidths
            Document doc = new Document(MyDir + "Tables.docx");

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

        [Test]
        public static void AutoFitTableToPageWidth()
        {
            // ExStart:AutoFitTableToPageWidth
            Document doc = new Document(MyDir + "Tables.docx");

            Table table = (Table) doc.GetChild(NodeType.Table, 0, true);
            // Autofit the first table to the page width
            table.AutoFit(AutoFitBehavior.AutoFitToWindow);

            doc.Save(ArtifactsDir + "AutoFitTableToWindow.docx");

            Debug.Assert(doc.FirstSection.Body.Tables[0].PreferredWidth.Type == PreferredWidthType.Percent,
                "PreferredWidth type is not percent");
            Debug.Assert(doc.FirstSection.Body.Tables[0].PreferredWidth.Value == 100,
                "PreferredWidth value is different than 100");
            //ExEnd:AutoFitTableToPageWidth
        }

        [Test]
        public static void BuildTableFromDataTable()
        {
            //ExStart:BuildTableFromDataTable
            Document doc = new Document();

            // We can position where we want the table to be inserted and also specify any extra formatting to be
            // applied onto the table as well
            DocumentBuilder builder = new DocumentBuilder(doc);

            // We want to rotate the page landscape as we expect a wide table
            doc.FirstSection.PageSetup.Orientation = Orientation.Landscape;

            DataSet ds = new DataSet();
            ds.ReadXml(MyDir + "Employees.xml");
            // Retrieve the data from our data source which is stored as a DataTable
            DataTable dataTable = ds.Tables[0];

            // Build a table in the document from the data contained in the DataTable
            Table table = ImportTableFromDataTable(builder, dataTable, true);

            // We can apply a table style as a very quick way to apply formatting to the entire table
            table.StyleIdentifier = StyleIdentifier.MediumList2Accent1;
            table.StyleOptions = TableStyleOptions.FirstRow | TableStyleOptions.RowBands | TableStyleOptions.LastColumn;

            // For our table we want to remove the heading for the image column
            table.FirstRow.LastCell.RemoveAllChildren();

            doc.Save(ArtifactsDir + "BuildTableFromDataTable.docx");
            //ExEnd:BuildTableFromDataTable
        }

        //ExStart:ImportTableFromDataTable
        /// <summary>
        /// Imports the content from the specified DataTable into a new Aspose.Words Table object. 
        /// The table is inserted at the current position of the document builder and using the current builder's formatting if any is defined.
        /// </summary>
        public static Table ImportTableFromDataTable(DocumentBuilder builder, DataTable dataTable,
            bool importColumnHeadings)
        {
            Table table = builder.StartTable();

            // Check if the names of the columns from the data source are to be included in a header row
            if (importColumnHeadings)
            {
                // Store the original values of these properties before changing them
                bool boldValue = builder.Font.Bold;
                ParagraphAlignment paragraphAlignmentValue = builder.ParagraphFormat.Alignment;

                // Format the heading row with the appropriate properties
                builder.Font.Bold = true;
                builder.ParagraphFormat.Alignment = ParagraphAlignment.Center;

                // Create a new row and insert the name of each column into the first row of the table
                foreach (DataColumn column in dataTable.Columns)
                {
                    builder.InsertCell();
                    builder.Writeln(column.ColumnName);
                }

                builder.EndRow();

                // Restore the original formatting
                builder.Font.Bold = boldValue;
                builder.ParagraphFormat.Alignment = paragraphAlignmentValue;
            }

            foreach (DataRow dataRow in dataTable.Rows)
            {
                foreach (object item in dataRow.ItemArray)
                {
                    // Insert a new cell for each object
                    builder.InsertCell();

                    switch (item.GetType().Name)
                    {
                        case "DateTime":
                            // Define a custom format for dates and times
                            DateTime dateTime = (DateTime) item;
                            builder.Write(dateTime.ToString("MMMM d, yyyy"));
                            break;
                        default:
                            // By default any other item will be inserted as text
                            builder.Write(item.ToString());
                            break;
                    }
                }

                // After we insert all the data from the current record we can end the table row
                builder.EndRow();
            }

            // We have finished inserting all the data from the DataTable, we can end the table
            builder.EndTable();

            return table;
        }
        //ExEnd:ImportTableFromDataTable

        /// <summary>
        /// Shows how to clone complete table.
        /// </summary>
        [Test]
        public static void CloneCompleteTable()
        {
            //ExStart:CloneCompleteTable
            Document doc = new Document(MyDir + "Tables.docx");

            // Retrieve the first table in the document
            Table table = (Table) doc.GetChild(NodeType.Table, 0, true);

            // Create a clone of the table
            Table tableClone = (Table) table.Clone(true);

            // Insert the cloned table into the document after the original
            table.ParentNode.InsertAfter(tableClone, table);

            // Insert an empty paragraph between the two tables or else they will be combined into one upon save
            // This has to do with document validation
            table.ParentNode.InsertAfter(new Paragraph(doc), table);
            
            doc.Save(ArtifactsDir + "CloneCompleteTable.docx");
            //ExEnd:CloneCompleteTable
        }

        /// <summary>
        /// Shows how to clone last row of table.
        /// </summary>
        [Test]
        public static void CloneLastRow()
        {
            //ExStart:CloneLastRow
            Document doc = new Document(MyDir + "Tables.docx");

            // Retrieve the first table in the document
            Table table = (Table) doc.GetChild(NodeType.Table, 0, true);

            // Clone the last row in the table
            Row clonedRow = (Row) table.LastRow.Clone(true);

            // Remove all content from the cloned row's cells
            // This makes the row ready for new content to be inserted into
            foreach (Cell cell in clonedRow.Cells)
                cell.RemoveAllChildren();

            // Add the row to the end of the table.
            table.AppendChild(clonedRow);

            doc.Save(ArtifactsDir + "CloneLastRow.docx");
            //ExEnd:CloneLastRow
        }
        
        [Test]
        public static void ReplaceText()
        {
            //ExStart:ReplaceText
            Document doc = new Document(MyDir + "Tables.docx");

            // Get the first table in the document
            Table table = (Table) doc.GetChild(NodeType.Table, 0, true);
            // Replace any instances of our string in the entire table
            table.Range.Replace("Carrots", "Eggs", new FindReplaceOptions(FindReplaceDirection.Forward));
            // Replace any instances of our string in the last cell of the table only
            table.LastRow.LastCell.Range.Replace("50", "20", new FindReplaceOptions(FindReplaceDirection.Forward));

            doc.Save(ArtifactsDir + "ReplaceText.docx");
            //ExEnd:ReplaceText
        }

        [Test]
        public static void FindingIndex()
        {
            Document doc = new Document(MyDir + "Tables.docx");

            //ExStart:RetrieveTableIndex
            // Get the first table in the document
            Table table = (Table) doc.GetChild(NodeType.Table, 0, true);

            NodeCollection allTables = doc.GetChildNodes(NodeType.Table, true);
            int tableIndex = allTables.IndexOf(table);
            //ExEnd:RetrieveTableIndex
            Console.WriteLine("\nTable index is " + tableIndex);

            //ExStart:RetrieveRowIndex
            int rowIndex = table.IndexOf(table.LastRow);
            //ExEnd:RetrieveRowIndex
            Console.WriteLine("\nRow index is " + rowIndex);

            Row row = table.LastRow;
            //ExStart:RetrieveCellIndex
            int cellIndex = row.IndexOf(row.Cells[4]);
            //ExEnd:RetrieveCellIndex
            Console.WriteLine("\nCell index is " + cellIndex);
        }

        [Test]
        public static void InsertTableDirectly()
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

        [Test]
        public static void InsertTableFromHtml()
        {
            //ExStart:InsertTableFromHtml
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Insert the table from HTML. Note that AutoFitSettings does not apply to tables
            // inserted from HTML
            builder.InsertHtml("<table>" +
                               "<tr>" +
                               "<td>Row 1, Cell 1</td>" +
                               "<td>Row 1, Cell 2</td>" +
                               "</tr>" +
                               "<tr>" +
                               "<td>Row 2, Cell 2</td>" +
                               "<td>Row 2, Cell 2</td>" +
                               "</tr>" +
                               "</table>");

            doc.Save(ArtifactsDir + "InsertTableFromHtml.docx");
            //ExEnd:InsertTableFromHtml
        }

        [Test]
        public static void SimpleTable()
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

        [Test]
        public static void FormattedTable()
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

        [Test]
        public static void NestedTable()
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

        /// <summary>
        /// Shows how to combine the rows from two tables into one.
        /// </summary>
        [Test]
        public static void CombineRows()
        {
            //ExStart:CombineRows
            Document doc = new Document(MyDir + "Tables.docx");

            // Get the first and second table in the document
            // The rows from the second table will be appended to the end of the first table
            Table firstTable = (Table) doc.GetChild(NodeType.Table, 0, true);
            Table secondTable = (Table) doc.GetChild(NodeType.Table, 1, true);

            // Append all rows from the current table to the next
            // Due to the design of tables even tables with different cell count and widths can be joined into one table
            while (secondTable.HasChildNodes)
                firstTable.Rows.Add(secondTable.FirstRow);

            // Remove the empty table container
            secondTable.Remove();

            doc.Save(ArtifactsDir + "CombineRows.docx");
            //ExEnd:CombineRows
        }

        /// <summary>
        /// Shows how to split a table into two tables in a specific row.
        /// </summary>
        [Test]
        public static void SplitTable()
        {
            //ExStart:SplitTable
            Document doc = new Document(MyDir + "Tables.docx");

            // Get the first table in the document
            Table firstTable = (Table) doc.GetChild(NodeType.Table, 0, true);

            // We will split the table at the third row (inclusive)
            Row row = firstTable.Rows[2];

            // Create a new container for the split table
            Table table = (Table) firstTable.Clone(false);

            // Insert the container after the original
            firstTable.ParentNode.InsertAfter(table, firstTable);

            // Add a buffer paragraph to ensure the tables stay apart
            firstTable.ParentNode.InsertAfter(new Paragraph(doc), firstTable);

            Row currentRow;

            do
            {
                currentRow = firstTable.LastRow;
                table.PrependChild(currentRow);
            } while (currentRow != row);

            doc.Save(ArtifactsDir + "SplitTable.docx");
            //ExEnd:SplitTable
        }

        [Test]
        public static void RowFormatDisableBreakAcrossPages()
        {
            //ExStart:RowFormatDisableBreakAcrossPages
            Document doc = new Document(MyDir + "Table spanning two pages.docx");

            // Retrieve the first table in the document
            Table table = (Table) doc.GetChild(NodeType.Table, 0, true);

            // Disable breaking across pages for all rows in the table
            foreach (Row row in table.Rows)
                row.RowFormat.AllowBreakAcrossPages = false;

            doc.Save(ArtifactsDir + "RowFormatDisableBreakAcrossPages.docx");
            //ExEnd:RowFormatDisableBreakAcrossPages
        }

        [Test]
        public static void KeepTableTogether()
        {
            //ExStart:KeepTableTogether
            Document doc = new Document(MyDir + "Table spanning two pages.docx");
            // Retrieve the first table in the document
            Table table = (Table) doc.GetChild(NodeType.Table, 0, true);

            // To keep a table from breaking across a page we need to enable KeepWithNext 
            // for every paragraph in the table except for the last paragraphs in the last 
            // row of the table
            foreach (Cell cell in table.GetChildNodes(NodeType.Cell, true))
            {
                // Call this method if table's cell is created on the fly
                // Newly created cell does not have paragraph inside
                cell.EnsureMinimum();
                foreach (Paragraph para in cell.Paragraphs)
                    if (!(cell.ParentRow.IsLastRow && para.IsEndOfCell))
                        para.ParagraphFormat.KeepWithNext = true;
            }

            doc.Save(ArtifactsDir + "KeepTableTogether.docx");
            //ExEnd:KeepTableTogether
        }

        public static void Run()
        {
            CheckCellsMerged();
            // The below method shows how to create a table with two rows with cells in the first row horizontally merged
            HorizontalMerge();
            // The below method shows how to create a table with two columns with cells merged vertically in the first column
            VerticalMerge();
            // The below method shows how to merges the range of cells between the two specified cells
            MergeCellRange();
            // Show how to prints the horizontal and vertical merge of a cell
            PrintHorizontalAndVerticalMerged();
            // This method converts cells which are horizontally merged by its width to the cell horizontally merged by flags
            ConvertToHorizontallyMergedCells();
        }

        [Test]
        public static void CheckCellsMerged()
        {
            //ExStart:CheckCellsMerged 
            Document doc = new Document(MyDir + "Table with merged cells.docx");

            // Retrieve the first table in the document
            Table table = (Table) doc.GetChild(NodeType.Table, 0, true);

            foreach (Row row in table.Rows)
            {
                foreach (Cell cell in row.Cells)
                {
                    Console.WriteLine(PrintCellMergeType(cell));
                }
            }
            //ExEnd:CheckCellsMerged 
        }

        //ExStart:PrintCellMergeType 
        public static string PrintCellMergeType(Cell cell)
        {
            bool isHorizontallyMerged = cell.CellFormat.HorizontalMerge != CellMerge.None;
            bool isVerticallyMerged = cell.CellFormat.VerticalMerge != CellMerge.None;
            
            string cellLocation =
                $"R{cell.ParentRow.ParentTable.IndexOf(cell.ParentRow) + 1}, C{cell.ParentRow.IndexOf(cell) + 1}";

            if (isHorizontallyMerged && isVerticallyMerged)
                return $"The cell at {cellLocation} is both horizontally and vertically merged";
            
            if (isHorizontallyMerged)
                return $"The cell at {cellLocation} is horizontally merged.";
            
            if (isVerticallyMerged)
                return $"The cell at {cellLocation} is vertically merged";
            
            return $"The cell at {cellLocation} is not merged";
        }
        //ExEnd:PrintCellMergeType
        
        [Test]
        public static void VerticalMerge()
        {
            //ExStart:VerticalMerge           
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            builder.InsertCell();
            builder.CellFormat.VerticalMerge = CellMerge.First;
            builder.Write("Text in merged cells.");

            builder.InsertCell();
            builder.CellFormat.VerticalMerge = CellMerge.None;
            builder.Write("Text in one cell");
            builder.EndRow();

            builder.InsertCell();
            // This cell is vertically merged to the cell above and should be empty
            builder.CellFormat.VerticalMerge = CellMerge.Previous;

            builder.InsertCell();
            builder.CellFormat.VerticalMerge = CellMerge.None;
            builder.Write("Text in another cell");
            builder.EndRow();
            builder.EndTable();
            
            doc.Save(ArtifactsDir + "VerticalMerge.docx");
            //ExEnd:VerticalMerge
        }

        [Test]
        public static void HorizontalMerge()
        {
            //ExStart:HorizontalMerge         
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            builder.InsertCell();
            builder.CellFormat.HorizontalMerge = CellMerge.First;
            builder.Write("Text in merged cells.");

            builder.InsertCell();
            // This cell is merged to the previous and should be empty
            builder.CellFormat.HorizontalMerge = CellMerge.Previous;
            builder.EndRow();

            builder.InsertCell();
            builder.CellFormat.HorizontalMerge = CellMerge.None;
            builder.Write("Text in one cell.");

            builder.InsertCell();
            builder.Write("Text in another cell.");
            builder.EndRow();
            builder.EndTable();
            
            doc.Save(ArtifactsDir + "HorizontalMerge.docx");
            //ExEnd:HorizontalMerge
        }

        [Test]
        public static void MergeCellRange()
        {
            //ExStart:MergeCellRange
            Document doc = new Document(MyDir + "Table with merged cells.docx");

            // Retrieve the first table in the body of the first section
            Table table = doc.FirstSection.Body.Tables[0];

            // We want to merge the range of cells found inbetween these two cells
            Cell cellStartRange = table.Rows[0].Cells[0];
            Cell cellEndRange = table.Rows[1].Cells[1];

            // Merge all the cells between the two specified cells into one
            MergeCells(cellStartRange, cellEndRange);
            
            doc.Save(ArtifactsDir + "MergeCellRange.docx");
            //ExEnd:MergeCellRange
        }

        [Test]
        public static void PrintHorizontalAndVerticalMerged()
        {
            //ExStart:PrintHorizontalAndVerticalMerged
            Document doc = new Document(MyDir + "Table with merged cells.docx");

            // Create visitor
            SpanVisitor visitor = new SpanVisitor(doc);
            // Accept visitor
            doc.Accept(visitor);
            //ExEnd:PrintHorizontalAndVerticalMerged
        }

        [Test]
        public static void ConvertToHorizontallyMergedCells()
        {
            //ExStart:ConvertToHorizontallyMergedCells         
            Document doc = new Document(MyDir + "Table with merged cells.docx");

            Table table = doc.FirstSection.Body.Tables[0];
            table.ConvertToHorizontallyMergedCells(); // Now merged cells have appropriate merge flags
            //ExEnd:ConvertToHorizontallyMergedCells
        }

        //ExStart:MergeCells
        internal static void MergeCells(Cell startCell, Cell endCell)
        {
            Table parentTable = startCell.ParentRow.ParentTable;

            // Find the row and cell indices for the start and end cell
            Point startCellPos = new Point(startCell.ParentRow.IndexOf(startCell),
                parentTable.IndexOf(startCell.ParentRow));
            Point endCellPos = new Point(endCell.ParentRow.IndexOf(endCell), parentTable.IndexOf(endCell.ParentRow));
            // Create the range of cells to be merged based off these indices
            // Inverse each index if the end cell if before the start cell
            Rectangle mergeRange = new Rectangle(System.Math.Min(startCellPos.X, endCellPos.X),
                System.Math.Min(startCellPos.Y, endCellPos.Y),
                System.Math.Abs(endCellPos.X - startCellPos.X) + 1, System.Math.Abs(endCellPos.Y - startCellPos.Y) + 1);

            foreach (Row row in parentTable.Rows)
            {
                foreach (Cell cell in row.Cells)
                {
                    Point currentPos = new Point(row.IndexOf(cell), parentTable.IndexOf(row));

                    // Check if the current cell is inside our merge range then merge it
                    if (mergeRange.Contains(currentPos))
                    {
                        cell.CellFormat.HorizontalMerge = currentPos.X == mergeRange.X ? CellMerge.First : CellMerge.Previous;

                        cell.CellFormat.VerticalMerge = currentPos.Y == mergeRange.Y ? CellMerge.First : CellMerge.Previous;
                    }
                }
            }
        }
        //ExEnd:MergeCells
        
        //ExStart:HorizontalAndVerticalMergeHelperClasses
        /// <summary>
        /// Helper class that contains collection of rowinfo for each row
        /// </summary>
        public class TableInfo
        {
            public List<RowInfo> Rows { get; } = new List<RowInfo>();
        }

        /// <summary>
        /// Helper class that contains collection of cellinfo for each cell
        /// </summary>
        public class RowInfo
        {
            public List<CellInfo> Cells { get; } = new List<CellInfo>();
        }

        /// <summary>
        /// Helper class that contains info about cell. currently here is only colspan and rowspan
        /// </summary>
        public class CellInfo
        {
            public CellInfo(int colSpan, int rowSpan)
            {
                ColSpan = colSpan;
                RowSpan = rowSpan;
            }

            public int ColSpan { get; }
            public int RowSpan { get; }
        }

        public class SpanVisitor : DocumentVisitor
        {
            /// <summary>
            /// Creates new SpanVisitor instance.
            /// </summary>
            /// <param name="doc">
            /// Is document which we should parse.
            /// </param>
            public SpanVisitor(Document doc)
            {
                // Get collection of tables from the document
                mWordTables = doc.GetChildNodes(NodeType.Table, true);

                // Convert document to HTML
                // We will parse HTML to determine rowspan and colspan of each cell
                MemoryStream htmlStream = new MemoryStream();

                Saving.HtmlSaveOptions options = new Saving.HtmlSaveOptions();
                options.ImagesFolder = Path.GetTempPath();

                doc.Save(htmlStream, options);

                // Load HTML into the XML document
                XmlDocument xmlDoc = new XmlDocument();
                htmlStream.Position = 0;
                xmlDoc.Load(htmlStream);

                // Get collection of tables in the HTML document
                XmlNodeList tables = xmlDoc.DocumentElement.SelectNodes("// Table");

                foreach (XmlNode table in tables)
                {
                    TableInfo tableInf = new TableInfo();
                    // Get collection of rows in the table
                    XmlNodeList rows = table.SelectNodes("tr");

                    foreach (XmlNode row in rows)
                    {
                        RowInfo rowInf = new RowInfo();
                        // Get collection of cells
                        XmlNodeList cells = row.SelectNodes("td");

                        foreach (XmlNode cell in cells)
                        {
                            // Determine row span and colspan of the current cell
                            XmlAttribute colSpanAttr = cell.Attributes["colspan"];
                            XmlAttribute rowSpanAttr = cell.Attributes["rowspan"];

                            int colSpan = colSpanAttr == null ? 0 : int.Parse(colSpanAttr.Value);
                            int rowSpan = rowSpanAttr == null ? 0 : int.Parse(rowSpanAttr.Value);

                            CellInfo cellInf = new CellInfo(colSpan, rowSpan);
                            rowInf.Cells.Add(cellInf);
                        }

                        tableInf.Rows.Add(rowInf);
                    }

                    mTables.Add(tableInf);
                }
            }

            public override VisitorAction VisitCellStart(Cell cell)
            {
                // Determine index of current table
                int tabIdx = mWordTables.IndexOf(cell.ParentRow.ParentTable);

                // Determine index of current row
                int rowIdx = cell.ParentRow.ParentTable.IndexOf(cell.ParentRow);

                // And determine index of current cell
                int cellIdx = cell.ParentRow.IndexOf(cell);

                // Determine colspan and rowspan of current cell
                int colSpan = 0;
                int rowSpan = 0;
                if (tabIdx < mTables.Count &&
                    rowIdx < mTables[tabIdx].Rows.Count &&
                    cellIdx < mTables[tabIdx].Rows[rowIdx].Cells.Count)
                {
                    colSpan = mTables[tabIdx].Rows[rowIdx].Cells[cellIdx].ColSpan;
                    rowSpan = mTables[tabIdx].Rows[rowIdx].Cells[cellIdx].RowSpan;
                }

                Console.WriteLine("{0}.{1}.{2} colspan={3}\t rowspan={4}", tabIdx, rowIdx, cellIdx, colSpan, rowSpan);

                return VisitorAction.Continue;
            }

            private readonly List<TableInfo> mTables = new List<TableInfo>();
            private readonly NodeCollection mWordTables;
        }
        //ExEnd:HorizontalAndVerticalMergeHelperClasses

        [Test]
        public static void RepeatRowsOnSubsequentPages()
        {
            //ExStart:RepeatRowsOnSubsequentPages
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            builder.StartTable();
            builder.RowFormat.HeadingFormat = true;
            builder.ParagraphFormat.Alignment = ParagraphAlignment.Center;
            builder.CellFormat.Width = 100;
            builder.InsertCell();
            builder.Writeln("Heading row 1");
            builder.EndRow();
            builder.InsertCell();
            builder.Writeln("Heading row 2");
            builder.EndRow();

            builder.CellFormat.Width = 50;
            builder.ParagraphFormat.ClearFormatting();

            // Insert some content so the table is long enough to continue onto the next page
            for (int i = 0; i < 50; i++)
            {
                builder.InsertCell();
                builder.RowFormat.HeadingFormat = false;
                builder.Write("Column 1 Text");
                builder.InsertCell();
                builder.Write("Column 2 Text");
                builder.EndRow();
            }

            doc.Save(ArtifactsDir + "RepeatRowsOnSubsequentPages.docx");
            //ExEnd:RepeatRowsOnSubsequentPages
        }

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
            Document doc = new Document(MyDir + "Tables.docx");

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

        [Test]
        public static void GetTablePosition()
        {
            //ExStart:GetTablePosition
            Document doc = new Document(MyDir + "Tables.docx");

            // Retrieve the first table in the document
            Table table = (Table) doc.GetChild(NodeType.Table, 0, true);

            if (table.TextWrapping == TextWrapping.Around)
            {
                Console.WriteLine(table.RelativeHorizontalAlignment);
                Console.WriteLine(table.RelativeVerticalAlignment);
            }
            else
            {
                Console.WriteLine(table.Alignment);
            }
            //ExEnd:GetTablePosition
        }

        [Test]
        public static void GetFloatingTablePosition()
        {
            //ExStart:GetFloatingTablePosition
            Document doc = new Document(MyDir + "Floating table position.docx");
            
            foreach (Table table in doc.FirstSection.Body.Tables)
            {
                // If table is floating type then print its positioning properties
                if (table.TextWrapping == TextWrapping.Around)
                {
                    Console.WriteLine(table.HorizontalAnchor);
                    Console.WriteLine(table.VerticalAnchor);
                    Console.WriteLine(table.AbsoluteHorizontalDistance);
                    Console.WriteLine(table.AbsoluteVerticalDistance);
                    Console.WriteLine(table.AllowOverlap);
                    Console.WriteLine(table.AbsoluteHorizontalDistance);
                    Console.WriteLine(table.RelativeVerticalAlignment);
                    Console.WriteLine("..............................");
                }
            }
            //ExEnd:GetFloatingTablePosition
        }

        [Test]
        public static void SetFloatingTablePosition()
        {
            //ExStart:SetFloatingTablePosition
            Document doc = new Document(MyDir + "Floating table position.docx");

            Table table = doc.FirstSection.Body.Tables[0];
            // Sets absolute table horizontal position at 10pt
            table.AbsoluteHorizontalDistance = 10;
            // Sets vertical table position to center of entity specified by Table.VerticalAnchor
            table.RelativeVerticalAlignment = VerticalAlignment.Center;

            doc.Save(ArtifactsDir + "SetFloatingTablePosition.docx");
            //ExEnd:SetFloatingTablePosition
        }

        [Test]
        public static void SetRelativeHorizontalOrVerticalPosition()
        {
            // ExStart:SetRelativeHorizontalOrVerticalPosition
            Document doc = new Document(MyDir + "Floating table position.docx");
            Table table = doc.FirstSection.Body.Tables[0];

            table.HorizontalAnchor = RelativeHorizontalPosition.Column;
            table.VerticalAnchor = RelativeVerticalPosition.Page;

            // Save the document to disk.
            doc.Save(ArtifactsDir + "Table.SetFloatingTablePosition.docx");
            // ExEnd:SetRelativeHorizontalOrVerticalPosition
        }
    }
}