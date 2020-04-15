using System.Text;
using System.Collections;
using System;
using Aspose.Words.Tables;
using NUnit.Framework;

namespace Aspose.Words.Examples.CSharp.Programming_Documents.Working_With_Document
{
    class AddRemoveColumn : TestDataHelper
    {
        [Test]
        public static void RemoveColumn()
        {
            //ExStart:RemoveColumn
            Document doc = new Document(TablesDir + "Table.Document.doc");

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
            Document doc = new Document(TablesDir + "Table.Document.doc");

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
    }
}