using System;
using Aspose.Words.Tables;

namespace Aspose.Words.Examples.CSharp.Programming_Documents.Working_With_Document
{
    class FindingIndex : TestDataHelper
    {
        public static void Run()
        {
            Document doc = new Document(TablesDir + "Table.SimpleTable.doc");

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
    }
}