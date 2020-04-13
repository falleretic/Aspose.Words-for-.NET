using Aspose.Words.Tables;

namespace Aspose.Words.Examples.CSharp.Programming_Documents.Working_with_Tables
{
    class JoiningAndSplittingTable : TestDataHelper
    {
        public static void Run()
        {
            CombineRows();
            SplitTable();
        }

        /// <summary>
        /// Shows how to combine the rows from two tables into one.
        /// </summary>        
        private static void CombineRows()
        {
            //ExStart:CombineRows
            Document doc = new Document(TablesDir + "Table.Document.doc");

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
        private static void SplitTable()
        {
            //ExStart:SplitTable
            Document doc = new Document(TablesDir + "Table.Document.doc");

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
    }
}