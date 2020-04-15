using NUnit.Framework;

namespace Aspose.Words.Examples.CSharp.Programming_Documents.Bookmarks
{
    class AccessBookmarks : TestDataHelper
    {
        [Test]
        public static void Run()
        {
            //ExStart:AccessBookmarks
            Document doc = new Document(BookmarksDir + "Bookmarks.doc");
            
            // By index
            Bookmark bookmark1 = doc.Range.Bookmarks[0];
            
            // By name
            Bookmark bookmark2 = doc.Range.Bookmarks["Bookmark2"];
            //ExEnd:AccessBookmarks
        }
    }
}