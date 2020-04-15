using System;
using Aspose.Words.Tables;
using NUnit.Framework;

namespace Aspose.Words.Examples.CSharp.Programming_Documents.Bookmarks
{
    class UntangleRowBookmarks : TestDataHelper
    {
        [Test]
        public static void Run()
        {
            Document doc = new Document(BookmarksDir + "TestDefect1352.doc");

            // This perform the custom task of putting the row bookmark ends into the same row with the bookmark starts
            Untangle(doc);

            // Now we can easily delete rows by a bookmark without damaging any other row's bookmarks
            DeleteRowByBookmark(doc, "ROW2");

            // This is just to check that the other bookmark was not damaged
            if (doc.Range.Bookmarks["ROW1"].BookmarkEnd == null)
                throw new Exception("Wrong, the end of the bookmark was deleted.");

            doc.Save(ArtifactsDir + "TestDefect1352.doc");
        }

        private static void Untangle(Document doc)
        {
            foreach (Bookmark bookmark in doc.Range.Bookmarks)
            {
                // Get the parent row of both the bookmark and bookmark end node
                Row row1 = (Row) bookmark.BookmarkStart.GetAncestor(typeof(Row));
                Row row2 = (Row) bookmark.BookmarkEnd.GetAncestor(typeof(Row));

                // If both rows are found okay and the bookmark start and end are contained
                // In adjacent rows, then just move the bookmark end node to the end
                // Of the last paragraph in the last cell of the top row
                if (row1 != null && row2 != null && row1.NextSibling == row2)
                    row1.LastCell.LastParagraph.AppendChild(bookmark.BookmarkEnd);
            }
        }

        private static void DeleteRowByBookmark(Document doc, string bookmarkName)
        {
            // Find the bookmark in the document
            Bookmark bookmark = doc.Range.Bookmarks[bookmarkName];

            // Get the parent row of the bookmark
            Row row = (Row) bookmark?.BookmarkStart.GetAncestor(typeof(Row));

            // Remove the row
            row?.Remove();
        }
    }
}