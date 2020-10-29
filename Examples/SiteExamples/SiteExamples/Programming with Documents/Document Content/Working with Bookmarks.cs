using System;
using Aspose.Words;
using Aspose.Words.Fields;
using Aspose.Words.Saving;
using Aspose.Words.Tables;
using NUnit.Framework;

namespace SiteExamples.Programming_with_Documents.Document_Content
{
    class BookmarksExamples : SiteExamplesBase
    {
        [Test]
        public static void AccessBookmarks()
        {
            //ExStart:AccessBookmarks
            Document doc = new Document(MyDir + "Bookmarks.docx");
            
            // By index
            Bookmark bookmark1 = doc.Range.Bookmarks[0];
            // By name
            Bookmark bookmark2 = doc.Range.Bookmarks["MyBookmark3"];
            //ExEnd:AccessBookmarks
        }

        [Test]
        public static void BookmarkNameAndText()
        {
            //ExStart:BookmarkNameAndText
            Document doc = new Document(MyDir + "Bookmarks.docx");

            // Use the indexer of the Bookmarks collection to obtain the desired bookmark
            Bookmark bookmark = doc.Range.Bookmarks["MyBookmark1"];

            // Get the name and text of the bookmark
            string name = bookmark.Name;
            string text = bookmark.Text;

            // Set the name and text of the bookmark
            bookmark.Name = "RenamedBookmark";
            bookmark.Text = "This is a new bookmarked text.";
            //ExEnd:BookmarkNameAndText
        }

        [Test]
        public static void BookmarkTableColumns()
        {
            //ExStart:BookmarkTable
            // Create empty document
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            builder.StartTable();

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
            //ExEnd:BookmarkTable

            //ExStart:BookmarkTableColumns
            foreach (Bookmark bookmark in doc.Range.Bookmarks)
            {
                Console.WriteLine("Bookmark: {0}{1}", bookmark.Name, bookmark.IsColumn ? " (Column)" : "");

                if (bookmark.IsColumn)
                {
                    if (bookmark.BookmarkStart.GetAncestor(NodeType.Row) is Row row && bookmark.FirstColumn < row.Cells.Count)
                        Console.WriteLine(row.Cells[bookmark.FirstColumn].GetText().TrimEnd(ControlChar.CellChar));
                }
            }
            //ExEnd:BookmarkTableColumns
        }

        [Test]
        public static void CopyBookmarkedText()
        {
            Document srcDoc = new Document(MyDir + "Bookmarks.docx");

            // This is the bookmark whose content we want to copy
            Bookmark srcBookmark = srcDoc.Range.Bookmarks["MyBookmark1"];

            // We will be adding to this document
            Document dstDoc = new Document();

            // Let's say we will be appending to the end of the body of the last section
            CompositeNode dstNode = dstDoc.LastSection.Body;

            // It is a good idea to use this import context object because multiple nodes are being imported
            // If you import multiple times without a single context, it will result in many styles created
            NodeImporter importer = new NodeImporter(srcDoc, dstDoc, ImportFormatMode.KeepSourceFormatting);

            // Do it once
            AppendBookmarkedText(importer, srcBookmark, dstNode);

            // Do it one more time for fun
            AppendBookmarkedText(importer, srcBookmark, dstNode);

            dstDoc.Save(ArtifactsDir + "Template.docx");
        }

        /// <summary>
        /// Copies content of the bookmark and adds it to the end of the specified node.
        /// The destination node can be in a different document.
        /// </summary>
        /// <param name="importer">Maintains the import context </param>
        /// <param name="srcBookmark">The input bookmark</param>
        /// <param name="dstNode">Must be a node that can contain paragraphs (such as a Story).</param>
        private static void AppendBookmarkedText(NodeImporter importer, Bookmark srcBookmark, CompositeNode dstNode)
        {
            // This is the paragraph that contains the beginning of the bookmark.
            Paragraph startPara = (Paragraph) srcBookmark.BookmarkStart.ParentNode;

            // This is the paragraph that contains the end of the bookmark.
            Paragraph endPara = (Paragraph) srcBookmark.BookmarkEnd.ParentNode;

            if (startPara == null || endPara == null)
                throw new InvalidOperationException(
                    "Parent of the bookmark start or end is not a paragraph, cannot handle this scenario yet.");

            // Limit ourselves to a reasonably simple scenario.
            if (startPara.ParentNode != endPara.ParentNode)
                throw new InvalidOperationException(
                    "Start and end paragraphs have different parents, cannot handle this scenario yet.");

            // We want to copy all paragraphs from the start paragraph up to (and including) the end paragraph,
            // Therefore the node at which we stop is one after the end paragraph.
            Node endNode = endPara.NextSibling;

            // This is the loop to go through all paragraph-level nodes in the bookmark.
            for (Node curNode = startPara; curNode != endNode; curNode = curNode.NextSibling)
            {
                // This creates a copy of the current node and imports it (makes it valid) in the context
                // Of the destination document. Importing means adjusting styles and list identifiers correctly.
                Node newNode = importer.ImportNode(curNode, true);

                // Now we simply append the new node to the destination.
                dstNode.AppendChild(newNode);
            }
        }

        [Test]
        public static void CreateBookmark()
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
        }

        [Test]
        public static void ShowHideBookmarks_call()
        {
            //ExStart:ShowHideBookmarks_call
            Document doc = new Document(MyDir + "Bookmarks.docx");

            ShowHideBookmarkedContent(doc, "MyBookmark1", false);
            
            doc.Save(ArtifactsDir + "UpdatedDocument.docx");
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

        [Test]
        public static void UntangleRowBookmarks()
        {
            Document doc = new Document(MyDir + "Table column bookmarks.docx");

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

        [Test]
        public static void DocumentBuilderInsertBookmark()
        {
            //ExStart:DocumentBuilderInsertBookmark
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            builder.StartBookmark("FineBookmark");
            builder.Writeln("This is just a fine bookmark.");
            builder.EndBookmark("FineBookmark");

            doc.Save(ArtifactsDir + "DocumentBuilderInsertBookmark.doc");
            //ExEnd:DocumentBuilderInsertBookmark
        }
    }
}
