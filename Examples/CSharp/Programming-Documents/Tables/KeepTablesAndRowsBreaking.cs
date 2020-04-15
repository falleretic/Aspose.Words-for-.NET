using Aspose.Words.Tables;
using NUnit.Framework;

namespace Aspose.Words.Examples.CSharp.Programming_Documents.Working_with_Tables
{
    class KeepTablesAndRowsBreaking : TestDataHelper
    {
        [Test]
        public static void RowFormatDisableBreakAcrossPages()
        {
            //ExStart:RowFormatDisableBreakAcrossPages
            Document doc = new Document(TablesDir + "Table.TableAcrossPage.doc");

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
            Document doc = new Document(TablesDir + "Table.TableAcrossPage.doc");
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
    }
}