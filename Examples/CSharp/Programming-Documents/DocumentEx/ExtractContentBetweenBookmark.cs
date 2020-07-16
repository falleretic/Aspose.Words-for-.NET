using System.Collections;
using NUnit.Framework;

namespace Aspose.Words.Examples.CSharp.DocumentEx
{
    class ExtractContentBetweenBookmark : TestDataHelper
    {
        [Test]
        public static void Run()
        {
            //ExStart:ExtractContentBetweenBookmark
            Document doc = new Document(DocumentDir + "TestFile.doc");

            Section section = doc.Sections[0];
            section.PageSetup.LeftMargin = 70.85;

            // Retrieve the bookmark from the document
            Bookmark bookmark = doc.Range.Bookmarks["Bookmark1"];

            // We use the BookmarkStart and BookmarkEnd nodes as markers
            BookmarkStart bookmarkStart = bookmark.BookmarkStart;
            BookmarkEnd bookmarkEnd = bookmark.BookmarkEnd;

            // Firstly extract the content between these nodes including the bookmark
            ArrayList extractedNodesInclusive = Common.ExtractContent(bookmarkStart, bookmarkEnd, true);
            Document dstDoc = Common.GenerateDocument(doc, extractedNodesInclusive);
            dstDoc.Save(ArtifactsDir + "TestFile.BookmarkInclusive.doc");

            // Secondly extract the content between these nodes this time without including the bookmark
            ArrayList extractedNodesExclusive = Common.ExtractContent(bookmarkStart, bookmarkEnd, false);
            dstDoc = Common.GenerateDocument(doc, extractedNodesExclusive);
            dstDoc.Save(ArtifactsDir + "TestFile.BookmarkExclusive.doc");
            //ExEnd:ExtractContentBetweenBookmark
        }
    }
}