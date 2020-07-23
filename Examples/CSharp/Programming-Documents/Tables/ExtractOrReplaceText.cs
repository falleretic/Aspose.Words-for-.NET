using System;
using Aspose.Words.Replacing;
using Aspose.Words.Tables;
using NUnit.Framework;

namespace Aspose.Words.Examples.CSharp.Tables
{
    class ExtractText : TestDataHelper
    {
        [Test]
        public static void ExtractPrintText()
        {
            //ExStart:ExtractText
            Document doc = new Document(TablesDir + "Tables.docx");

            // Get the first table in the document
            Table table = (Table) doc.GetChild(NodeType.Table, 0, true);

            // The range text will include control characters such as "\a" for a cell
            // You can call ToString and pass SaveFormat.Text on the desired node to find the plain text content

            // Print the plain text range of the table to the screen
            Console.WriteLine("Contents of the table: ");
            Console.WriteLine(table.Range.Text);
            //ExEnd:ExtractText   

            //ExStart:PrintTextRangeOFRowAndTable
            // Print the contents of the second row to the screen
            Console.WriteLine("\nContents of the row: ");
            Console.WriteLine(table.Rows[1].Range.Text);

            // Print the contents of the last cell in the table to the screen
            Console.WriteLine("\nContents of the cell: ");
            Console.WriteLine(table.LastRow.LastCell.Range.Text);
            //ExEnd:PrintTextRangeOFRowAndTable
        }

        [Test]
        public static void ReplaceText()
        {
            //ExStart:ReplaceText
            Document doc = new Document(TablesDir + "Tables.docx");

            // Get the first table in the document
            Table table = (Table) doc.GetChild(NodeType.Table, 0, true);
            // Replace any instances of our string in the entire table
            table.Range.Replace("Carrots", "Eggs", new FindReplaceOptions(FindReplaceDirection.Forward));
            // Replace any instances of our string in the last cell of the table only
            table.LastRow.LastCell.Range.Replace("50", "20", new FindReplaceOptions(FindReplaceDirection.Forward));

            doc.Save(ArtifactsDir + "ReplaceText.docx");
            //ExEnd:ReplaceText
        }
    }
}