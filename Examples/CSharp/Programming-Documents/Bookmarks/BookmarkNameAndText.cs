namespace Aspose.Words.Examples.CSharp.Programming_Documents.Bookmarks
{
    class BookmarkNameAndText : TestDataHelper
    {
        public static void Run()
        {
            //ExStart:BookmarkNameAndText
            Document doc = new Document(BookmarksDir + "Bookmark.doc");

            // Use the indexer of the Bookmarks collection to obtain the desired bookmark
            Bookmark bookmark = doc.Range.Bookmarks["MyBookmark"];

            // Get the name and text of the bookmark
            string name = bookmark.Name;
            string text = bookmark.Text;

            // Set the name and text of the bookmark
            bookmark.Name = "RenamedBookmark";
            bookmark.Text = "This is a new bookmarked text.";
            //ExEnd:BookmarkNameAndText
        }
    }
}