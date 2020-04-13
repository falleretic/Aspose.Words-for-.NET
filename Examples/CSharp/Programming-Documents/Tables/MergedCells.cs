using System;
using System.Collections.Generic;
using System.IO;
using System.Xml;
using System.Drawing;
using Aspose.Words.Tables;
using Aspose.Words.Saving;

namespace Aspose.Words.Examples.CSharp.Programming_Documents.Working_with_Tables
{
    class MergedCells : TestDataHelper
    {
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

        public static void CheckCellsMerged()
        {
            //ExStart:CheckCellsMerged 
            Document doc = new Document(TablesDir + "Table.MergedCells.doc");

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

        public static void MergeCellRange()
        {
            //ExStart:MergeCellRange
            Document doc = new Document(TablesDir + "Table.Document.doc");

            // Retrieve the first table in the body of the first section
            Table table = doc.FirstSection.Body.Tables[0];

            // We want to merge the range of cells found inbetween these two cells
            Cell cellStartRange = table.Rows[2].Cells[2];
            Cell cellEndRange = table.Rows[3].Cells[3];

            // Merge all the cells between the two specified cells into one
            MergeCells(cellStartRange, cellEndRange);
            
            doc.Save(ArtifactsDir + "MergeCellRange.docx");
            //ExEnd:MergeCellRange
        }

        public static void PrintHorizontalAndVerticalMerged()
        {
            //ExStart:PrintHorizontalAndVerticalMerged
            Document doc = new Document(TablesDir + "Table.MergedCells.doc");

            // Create visitor
            SpanVisitor visitor = new SpanVisitor(doc);
            // Accept visitor
            doc.Accept(visitor);
            //ExEnd:PrintHorizontalAndVerticalMerged
        }

        public static void ConvertToHorizontallyMergedCells()
        {
            //ExStart:ConvertToHorizontallyMergedCells         
            Document doc = new Document();

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

                HtmlSaveOptions options = new HtmlSaveOptions();
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
    }
}