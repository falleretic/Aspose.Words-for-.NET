using Aspose.Words.Fields;
using NUnit.Framework;

namespace Aspose.Words.Examples.CSharp.Programming_Documents.Bookmarks
{
    class ShowHideBookmarks : TestDataHelper
    {
        [Test]
        public static void Run()
        {
            //ExStart:ShowHideBookmarks_call
            Document doc = new Document(BookmarksDir + "Bookmarks.doc");

            ShowHideBookmarkedContent(doc, "Bookmark2", false);
            
            doc.Save(ArtifactsDir + "UpdatedDocument.doc");
            //ExEnd:ShowHideBookmarks_call
        }

        //ExStart:ShowHideBookmarks
        public static void ShowHideBookmarkedContent(Document doc, string bookmarkName, bool showHide)
        {
            DocumentBuilder builder = new DocumentBuilder(doc);
            Bookmark bm = doc.Range.Bookmarks[bookmarkName];

            builder.MoveToDocumentEnd();
            // {IF "{MERGEFIELD bookmark}" = "true" "" ""}
            Field field = builder.InsertField("IF \"", null);
            builder.MoveTo(field.Start.NextSibling);
            builder.InsertField("MERGEFIELD " + bookmarkName + "", null);
            builder.Write("\" = \"true\" ");
            builder.Write("\"");
            builder.Write("\"");
            builder.Write(" \"\"");

            Node currentNode = field.Start;
            bool flag = true;
            while (currentNode != null && flag)
            {
                if (currentNode.NodeType == NodeType.Run)
                    if (currentNode.ToString(SaveFormat.Text).Trim().Equals("\""))
                        flag = false;

                Node nextNode = currentNode.NextSibling;

                bm.BookmarkStart.ParentNode.InsertBefore(currentNode, bm.BookmarkStart);
                currentNode = nextNode;
            }

            Node endNode = bm.BookmarkEnd;
            flag = true;
            while (currentNode != null && flag)
            {
                if (currentNode.NodeType == NodeType.FieldEnd)
                    flag = false;

                Node nextNode = currentNode.NextSibling;

                bm.BookmarkEnd.ParentNode.InsertAfter(currentNode, endNode);
                endNode = currentNode;
                currentNode = nextNode;
            }

            doc.MailMerge.Execute(new string[] { bookmarkName }, new object[] { showHide });

            //MailMerge can be avoided by using the following
            //builder.MoveToMergeField(bookmarkName);
            //builder.Write(showHide ? "true" : "false");
        }
        //ExEnd:ShowHideBookmarks
    }
}