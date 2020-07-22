using System;
using NUnit.Framework;

namespace Aspose.Words.Examples.CSharp.DocumentEx
{
    class DocumentBuilderMovingCursor : TestDataHelper
    {
        [Test]
        public static void CursorPosition()
        {
            //ExStart:DocumentBuilderCursorPosition
            // Shows how to access the current node in a document builder
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            Node curNode = builder.CurrentNode;
            Paragraph curParagraph = builder.CurrentParagraph;
            //ExEnd:DocumentBuilderCursorPosition

            Console.WriteLine("\nCursor move to paragraph: " + curParagraph.GetText());
        }

        [Test]
        public static void MoveToNode()
        {
            //ExStart:DocumentBuilderMoveToNode
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            builder.MoveTo(doc.FirstSection.Body.LastParagraph);
            // ExEnd:DocumentBuilderMoveToNode
        }

        [Test]
        public static void MoveToDocumentStartEnd()
        {
            //ExStart:DocumentBuilderMoveToDocumentStartEnd
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            builder.MoveToDocumentEnd();
            Console.WriteLine("\nThis is the end of the document.");

            builder.MoveToDocumentStart();
            Console.WriteLine("\nThis is the beginning of the document.");
            //ExEnd:DocumentBuilderMoveToDocumentStartEnd            
        }

        [Test]
        public static void MoveToSection()
        {
            //ExStart:DocumentBuilderMoveToSection
            // Create a blank document and append a section to it, giving it two sections
            Document doc = new Document();
            doc.AppendChild(new Section(doc));

            DocumentBuilder builder = new DocumentBuilder(doc);

            // Parameters are 0-index. Moves to third section
            builder.MoveToSection(1);
            builder.Writeln("This is the 2rd section.");
            //ExEnd:DocumentBuilderMoveToSection               
        }

        [Test]
        public static void HeadersAndFooters()
        {
            //ExStart:DocumentBuilderHeadersAndFooters
            // Create a blank document
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Specify that we want headers and footers different for first, even and odd pages
            builder.PageSetup.DifferentFirstPageHeaderFooter = true;
            builder.PageSetup.OddAndEvenPagesHeaderFooter = true;

            // Create the headers
            builder.MoveToHeaderFooter(HeaderFooterType.HeaderFirst);
            builder.Write("Header First");
            builder.MoveToHeaderFooter(HeaderFooterType.HeaderEven);
            builder.Write("Header Even");
            builder.MoveToHeaderFooter(HeaderFooterType.HeaderPrimary);
            builder.Write("Header Odd");

            // Create three pages in the document
            builder.MoveToSection(0);
            builder.Writeln("Page1");
            builder.InsertBreak(BreakType.PageBreak);
            builder.Writeln("Page2");
            builder.InsertBreak(BreakType.PageBreak);
            builder.Writeln("Page3");

            doc.Save(ArtifactsDir + "DocumentBuilder.HeadersAndFooters.doc");
            //ExEnd:DocumentBuilderHeadersAndFooters
        }

        [Test]
        public static void MoveToParagraph()
        {
            //ExStart:DocumentBuilderMoveToParagraph
            Document doc = new Document(DocumentDir + "Paragraphs.docx");
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Parameters are 0-index. Moves to third paragraph
            builder.MoveToParagraph(2, 0);
            builder.Writeln("This is the 3rd paragraph.");
            //ExEnd:DocumentBuilderMoveToParagraph               
        }

        [Test]
        public static void MoveToTableCell()
        {
            //ExStart:DocumentBuilderMoveToTableCell
            Document doc = new Document(DocumentDir + "Tables.docx");
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Move the builder to row 3, cell 4 of the first table
            builder.MoveToCell(0, 2, 3, 0);
            builder.Writeln("Hello World!");
            //ExEnd:DocumentBuilderMoveToTableCell               
        }

        [Test]
        public static void MoveToBookmark()
        {
            //ExStart:DocumentBuilderMoveToBookmark
            Document doc = new Document(DocumentDir + "Bookmarks.docx");
            DocumentBuilder builder = new DocumentBuilder(doc);

            builder.MoveToBookmark("MyBookmark1");
            builder.Writeln("This is a very cool bookmark.");
            //ExEnd:DocumentBuilderMoveToBookmark               
        }

        [Test]
        public static void MoveToBookmarkEnd()
        {
            //ExStart:DocumentBuilderMoveToBookmarkEnd
            Document doc = new Document(DocumentDir + "Bookmarks.docx");
            DocumentBuilder builder = new DocumentBuilder(doc);

            builder.MoveToBookmark("MyBookmark1", false, true);
            builder.Writeln("This is a very cool bookmark.");
            //ExEnd:DocumentBuilderMoveToBookmarkEnd              
        }

        [Test]
        public static void MoveToMergeField()
        {
            //ExStart:DocumentBuilderMoveToMergeField
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);
            
            builder.InsertField(@"MERGEFIELD MyMergeField1 \* MERGEFORMAT");
            builder.InsertField(@"MERGEFIELD MyMergeField2 \* MERGEFORMAT");

            builder.MoveToMergeField("MyMergeField1");
            builder.Writeln("This is a very nice merge field.");
            //ExEnd:DocumentBuilderMoveToMergeField              
        }
    }
}