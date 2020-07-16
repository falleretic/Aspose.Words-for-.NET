using Aspose.Words.Tables;
using NUnit.Framework;

namespace Aspose.Words.Examples.CSharp.Tables
{
    class CloneTable : TestDataHelper
    {
        /// <summary>
        /// Shows how to clone complete table.
        /// </summary>
        [Test]
        public static void CloneCompleteTable()
        {
            //ExStart:CloneCompleteTable
            Document doc = new Document(TablesDir + "Table.SimpleTable.doc");

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
            Document doc = new Document(TablesDir + "Table.SimpleTable.doc");

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
    }
}