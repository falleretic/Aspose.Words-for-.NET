using System;

namespace Aspose.Words.Examples.CSharp.Programming_Documents.Bookmarks
{
    class AccessBookmarks : TestDataHelper
    {
        public static void Run()
        {
            //ExStart:AccessBookmarks
            Document doc = new Document(BookmarksDir + "Bookmarks.doc");
            
            // By index
            Bookmark bookmark1 = doc.Range.Bookmarks[0];
            
            // By name
            Bookmark bookmark2 = doc.Range.Bookmarks["Bookmark2"];
            //ExEnd:AccessBookmarks

            Console.WriteLine("\nBookmark by name is " + bookmark1.Name + " and bookmark by index is " + bookmark2.Name);
        }
    }
}