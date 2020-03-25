using System;
using Aspose.Words.Tables;

namespace Aspose.Words.Examples.CSharp.Programming_Documents.Bookmarks
{
    class BookmarkTable : TestDataHelper
    {
        public static void Run()
        {
            InsertBookmarkTable();
            BookmarkTableColumns();
        }

        public static void InsertBookmarkTable()
        {
            //ExStart:BookmarkTable
            // Create empty document
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            Table table = builder.StartTable();

            // Insert a cell
            builder.InsertCell();

            // Start bookmark here after calling InsertCell
            builder.StartBookmark("MyBookmark");

            builder.Write("This is row 1 cell 1");

            // Insert a cell
            builder.InsertCell();
            builder.Write("This is row 1 cell 2");

            builder.EndRow();

            // Insert a cell
            builder.InsertCell();
            builder.Writeln("This is row 2 cell 1");

            // Insert a cell
            builder.InsertCell();
            builder.Writeln("This is row 2 cell 2");

            builder.EndRow();

            builder.EndTable();
            // End of bookmark
            builder.EndBookmark("MyBookmark");

            doc.Save(ArtifactsDir + "Bookmark.Table.doc");
            //ExEnd:BookmarkTable

            Console.WriteLine("\nTable bookmarked successfully.");
        }

        public static void BookmarkTableColumns()
        {
            //ExStart:BookmarkTableColumns
            Document doc = new Document(ArtifactsDir + "Bookmark.Table.doc");
            foreach (Bookmark bookmark in doc.Range.Bookmarks)
            {
                Console.WriteLine("Bookmark: {0}{1}", bookmark.Name, bookmark.IsColumn ? " (Column)" : "");

                if (bookmark.IsColumn)
                {
                    Row row = bookmark.BookmarkStart.GetAncestor(NodeType.Row) as Row;
                    if (row != null && bookmark.FirstColumn < row.Cells.Count)
                        Console.WriteLine(row.Cells[bookmark.FirstColumn].GetText().TrimEnd(ControlChar.CellChar));
                }
            }
            //ExEnd:BookmarkTableColumns
        }
    }
}