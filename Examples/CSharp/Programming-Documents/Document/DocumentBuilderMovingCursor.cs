using System;
using NUnit.Framework;

namespace Aspose.Words.Examples.CSharp.Programming_Documents.Working_With_Document
{
    class DocumentBuilderMovingCursor : TestDataHelper
    {
        [Test]
        public static void CursorPosition()
        {
            //ExStart:DocumentBuilderCursorPosition
            // Shows how to access the current node in a document builder
            Document doc = new Document(DocumentDir + "DocumentBuilder.doc");
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
            Document doc = new Document(DocumentDir + "DocumentBuilder.doc");
            DocumentBuilder builder = new DocumentBuilder(doc);

            builder.MoveTo(doc.FirstSection.Body.LastParagraph);
            // ExEnd:DocumentBuilderMoveToNode
        }

        [Test]
        public static void MoveToDocumentStartEnd()
        {
            //ExStart:DocumentBuilderMoveToDocumentStartEnd
            Document doc = new Document(DocumentDir + "DocumentBuilder.doc");
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
            Document doc = new Document(DocumentDir + "DocumentBuilder.doc");
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Parameters are 0-index. Moves to third section
            builder.MoveToSection(2);
            builder.Writeln("This is the 3rd section.");
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
            Document doc = new Document(DocumentDir + "DocumentBuilder.doc");
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
            Document doc = new Document(DocumentDir + "DocumentBuilder.doc");
            DocumentBuilder builder = new DocumentBuilder(doc);

            // All parameters are 0-index. Moves to the 2nd table, 3rd row, 5th cell
            builder.MoveToCell(1, 2, 4, 0);
            builder.Writeln("Hello World!");
            //ExEnd:DocumentBuilderMoveToTableCell               
        }

        [Test]
        public static void MoveToBookmark()
        {
            //ExStart:DocumentBuilderMoveToBookmark
            Document doc = new Document(DocumentDir + "DocumentBuilder.doc");
            DocumentBuilder builder = new DocumentBuilder(doc);

            builder.MoveToBookmark("CoolBookmark");
            builder.Writeln("This is a very cool bookmark.");
            //ExEnd:DocumentBuilderMoveToBookmark               
        }

        [Test]
        public static void MoveToBookmarkEnd()
        {
            //ExStart:DocumentBuilderMoveToBookmarkEnd
            Document doc = new Document(DocumentDir + "DocumentBuilder.doc");
            DocumentBuilder builder = new DocumentBuilder(doc);

            builder.MoveToBookmark("CoolBookmark", false, true);
            builder.Writeln("This is a very cool bookmark.");
            //ExEnd:DocumentBuilderMoveToBookmarkEnd              
        }

        [Test]
        public static void MoveToMergeField()
        {
            //ExStart:DocumentBuilderMoveToMergeField
            Document doc = new Document(DocumentDir + "DocumentBuilder.doc");
            DocumentBuilder builder = new DocumentBuilder(doc);

            builder.MoveToMergeField("NiceMergeField");
            builder.Writeln("This is a very nice merge field.");
            //ExEnd:DocumentBuilderMoveToMergeField              
        }
    }
}