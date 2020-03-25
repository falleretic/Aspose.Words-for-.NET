using System;
using Aspose.Words.Saving;

namespace Aspose.Words.Examples.CSharp.Programming_Documents.Bookmarks
{
    class CreateBookmark : TestDataHelper
    {
        public static void Run()
        {
            //ExStart:CreateBookmark
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            builder.StartBookmark("My Bookmark");
            builder.Writeln("Text inside a bookmark.");

            builder.StartBookmark("Nested Bookmark");
            builder.Writeln("Text inside a NestedBookmark.");
            builder.EndBookmark("Nested Bookmark");

            builder.Writeln("Text after Nested Bookmark.");
            builder.EndBookmark("My Bookmark");

            PdfSaveOptions options = new PdfSaveOptions();
            options.OutlineOptions.BookmarksOutlineLevels.Add("My Bookmark", 1);
            options.OutlineOptions.BookmarksOutlineLevels.Add("Nested Bookmark", 2);

            doc.Save(ArtifactsDir + "Create.Bookmark.pdf", options);
            //ExEnd:CreateBookmark

            Console.WriteLine("\nBookmark created successfully.");
        }
    }
}