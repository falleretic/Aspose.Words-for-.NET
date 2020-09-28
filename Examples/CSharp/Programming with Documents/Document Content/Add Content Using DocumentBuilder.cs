﻿using System;
using System.Drawing;
using System.IO;
using System.Text.RegularExpressions;
using Aspose.Words.Drawing;
using Aspose.Words.Fields;
using Aspose.Words.Replacing;
using Aspose.Words.Tables;
using NUnit.Framework;

namespace Aspose.Words.Examples.CSharp.DocumentEx
{
    class AddContentUsingDocumentBuilder : TestDataHelper
    {
        [Test]
        public static void DocumentBuilderBuildTable()
        {
            //ExStart:DocumentBuilderBuildTable
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            Table table = builder.StartTable();

            // Insert a cell
            builder.InsertCell();
            // Use fixed column widths
            table.AutoFit(AutoFitBehavior.FixedColumnWidths);

            builder.CellFormat.VerticalAlignment = CellVerticalAlignment.Center;
            builder.Write("This is row 1 cell 1");

            // Insert a cell
            builder.InsertCell();
            builder.Write("This is row 1 cell 2");

            builder.EndRow();

            // Insert a cell
            builder.InsertCell();

            // Apply new row formatting
            builder.RowFormat.Height = 100;
            builder.RowFormat.HeightRule = HeightRule.Exactly;

            builder.CellFormat.Orientation = TextOrientation.Upward;
            builder.Writeln("This is row 2 cell 1");

            // Insert a cell
            builder.InsertCell();
            builder.CellFormat.Orientation = TextOrientation.Downward;
            builder.Writeln("This is row 2 cell 2");

            builder.EndRow();
            builder.EndTable();

            doc.Save(ArtifactsDir + "DocumentBuilderBuildTable.doc");
            //ExEnd:DocumentBuilderBuildTable
        }

        [Test]
        public static void DocumentBuilderInsertHorizontalRule()
        {
            //ExStart:DocumentBuilderInsertHorizontalRule
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            builder.Writeln("Insert a horizontal rule shape into the document.");
            builder.InsertHorizontalRule();

            doc.Save(ArtifactsDir + "DocumentBuilder.InsertHorizontalRule.doc");
            //ExEnd:DocumentBuilderInsertHorizontalRule
        }

        [Test]
        public static void DocumentBuilderHorizontalRuleFormat()
        {
            //ExStart:DocumentBuilderHorizontalRuleFormat
            DocumentBuilder builder = new DocumentBuilder();

            Shape shape = builder.InsertHorizontalRule();
            HorizontalRuleFormat horizontalRuleFormat = shape.HorizontalRuleFormat;

            horizontalRuleFormat.Alignment = HorizontalRuleAlignment.Center;
            horizontalRuleFormat.WidthPercent = 70;
            horizontalRuleFormat.Height = 3;
            horizontalRuleFormat.Color = Color.Blue;
            horizontalRuleFormat.NoShade = true;

            builder.Document.Save(ArtifactsDir + "HorizontalRuleFormat.docx");
            //ExEnd:DocumentBuilderHorizontalRuleFormat
        }

        [Test]
        public static void DocumentBuilderInsertBreak()
        {
            //ExStart:DocumentBuilderInsertBreak
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            builder.Writeln("This is page 1.");
            builder.InsertBreak(BreakType.PageBreak);

            builder.Writeln("This is page 2.");
            builder.InsertBreak(BreakType.PageBreak);

            builder.Writeln("This is page 3.");
            doc.Save(ArtifactsDir + "DocumentBuilderInsertBreak.doc");
            //ExEnd:DocumentBuilderInsertBreak
        }

        [Test]
        public static void InsertTextInputFormField()
        {
            //ExStart:DocumentBuilderInsertTextInputFormField
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);
            builder.InsertTextInput("TextInput", TextFormFieldType.Regular, "", "Hello", 0);

            doc.Save(ArtifactsDir + "DocumentBuilderInsertTextInputFormField.doc");
            //ExEnd:DocumentBuilderInsertTextInputFormField
        }

        [Test]
        public static void InsertCheckBoxFormField()
        {
            //ExStart:DocumentBuilderInsertCheckBoxFormField
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);
            builder.InsertCheckBox("CheckBox", true, true, 0);

            doc.Save(ArtifactsDir + "DocumentBuilderInsertCheckBoxFormField.doc");
            //ExEnd:DocumentBuilderInsertCheckBoxFormField
        }

        [Test]
        public static void InsertComboBoxFormField()
        {
            //ExStart:DocumentBuilderInsertComboBoxFormField
            string[] items = { "One", "Two", "Three" };

            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);
            builder.InsertComboBox("DropDown", items, 0);

            doc.Save(ArtifactsDir + "DocumentBuilderInsertComboBoxFormField.doc");
            //ExEnd:DocumentBuilderInsertComboBoxFormField
        }

        [Test]
        public static void InsertHtml()
        {
            //ExStart:DocumentBuilderInsertHtml
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);
            builder.InsertHtml(
                "<P align='right'>Paragraph right</P>" +
                "<b>Implicit paragraph left</b>" +
                "<div align='center'>Div center</div>" +
                "<h1 align='left'>Heading 1 left.</h1>");

            doc.Save(ArtifactsDir + "DocumentBuilderInsertHtml.doc");
            //ExEnd:DocumentBuilderInsertHtml
        }

        [Test]
        public static void InsertHyperlink()
        {
            //ExStart:DocumentBuilderInsertHyperlink
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);
            builder.Write("Please make sure to visit ");
            // Specify font formatting for the hyperlink
            builder.Font.Color = Color.Blue;
            builder.Font.Underline = Underline.Single;
            // Insert the link
            builder.InsertHyperlink("Aspose Website", "http://www.aspose.com", false);
            // Revert to default formatting
            builder.Font.ClearFormatting();
            builder.Write(" for more information.");

            doc.Save(ArtifactsDir + "DocumentBuilderInsertHyperlink.doc");
            //ExEnd:DocumentBuilderInsertHyperlink
        }

        [Test]
        public static void InsertTableOfContents()
        {
            //ExStart:DocumentBuilderInsertTableOfContents
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);
            // Insert a table of contents at the beginning of the document
            builder.InsertTableOfContents("\\o \"1-3\" \\h \\z \\u");
            // Start the actual document content on the second page
            builder.InsertBreak(BreakType.PageBreak);

            // Build a document with complex structure by applying different heading styles thus creating TOC entries
            builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Heading1;

            builder.Writeln("Heading 1");

            builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Heading2;

            builder.Writeln("Heading 1.1");
            builder.Writeln("Heading 1.2");

            builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Heading1;

            builder.Writeln("Heading 2");
            builder.Writeln("Heading 3");

            builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Heading2;

            builder.Writeln("Heading 3.1");

            builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Heading3;

            builder.Writeln("Heading 3.1.1");
            builder.Writeln("Heading 3.1.2");
            builder.Writeln("Heading 3.1.3");

            builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Heading2;

            builder.Writeln("Heading 3.2");
            builder.Writeln("Heading 3.3");

            // The newly inserted table of contents will be initially empty
            // It needs to be populated by updating the fields in the document
            //ExStart:UpdateFields
            doc.UpdateFields();
            //ExEnd:UpdateFields

            doc.Save(ArtifactsDir + "DocumentBuilderInsertTableOfContents.doc");
            //ExEnd:DocumentBuilderInsertTableOfContents
        }

        [Test]
        public static void InsertInlineImage()
        {
            //ExStart:DocumentBuilderInsertInlineImage
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            builder.InsertImage(DocumentDir + "Watermark.png");

            doc.Save(ArtifactsDir + "DocumentBuilderInsertInlineImage.doc");
            //ExEnd:DocumentBuilderInsertInlineImage
        }

        [Test]
        public static void InsertFloatingImage()
        {
            //ExStart:DocumentBuilderInsertFloatingImage
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            builder.InsertImage(DocumentDir + "Watermark.png",
                RelativeHorizontalPosition.Margin,
                100,
                RelativeVerticalPosition.Margin,
                100,
                200,
                100,
                WrapType.Square);

            doc.Save(ArtifactsDir + "DocumentBuilderInsertFloatingImage.doc");
            //ExEnd:DocumentBuilderInsertFloatingImage
        }

        [Test]
        public static void DocumentBuilderInsertParagraph()
        {
            //ExStart:DocumentBuilderInsertParagraph
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Specify font formatting
            Font font = builder.Font;
            font.Size = 16;
            font.Bold = true;
            font.Color = System.Drawing.Color.Blue;
            font.Name = "Arial";
            font.Underline = Underline.Dash;

            // Specify paragraph formatting
            ParagraphFormat paragraphFormat = builder.ParagraphFormat;
            paragraphFormat.FirstLineIndent = 8;
            paragraphFormat.Alignment = ParagraphAlignment.Justify;
            paragraphFormat.KeepTogether = true;

            builder.Writeln("A whole paragraph.");

            doc.Save(ArtifactsDir + "DocumentBuilderInsertParagraph.doc");
            // ExEnd:DocumentBuilderInsertParagraph
        }

        [Test]
        public static void DocumentBuilderInsertTCField()
        {
            //ExStart:DocumentBuilderInsertTCField
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Insert a TC field at the current document builder position
            builder.InsertField("TC \"Entry Text\" \\f t");

            doc.Save(ArtifactsDir + "DocumentBuilderInsertTCField.doc");
            //ExEnd:DocumentBuilderInsertTCField
        }

        [Test]
        public static void Run()
        {
            //ExStart:DocumentBuilderInsertTCFieldsAtText
            Document doc = new Document();

            FindReplaceOptions options = new FindReplaceOptions();
            // Highlight newly inserted content
            options.ApplyFont.HighlightColor = Color.DarkOrange;
            options.ReplacingCallback = new InsertTCFieldHandler("Chapter 1", "\\l 1");

            // Insert a TC field which displays "Chapter 1" just before the text "The Beginning" in the document
            doc.Range.Replace(new Regex("The Beginning"), "", options);
            //ExEnd:DocumentBuilderInsertTCFieldsAtText
        }


        //ExStart:InsertTCFieldHandler
        public sealed class InsertTCFieldHandler : IReplacingCallback
        {
            // Store the text and switches to be used for the TC fields
            private string mFieldText;
            private string mFieldSwitches;

            /// <summary>
            /// The switches to use for each TC field. Can be an empty string or null.
            /// </summary>
            public InsertTCFieldHandler(string switches)
                : this(string.Empty, switches)
            {
                mFieldSwitches = switches;
            }

            /// <summary>
            /// The display text and switches to use for each TC field. Display name can be an empty string or null.
            /// </summary>
            public InsertTCFieldHandler(string text, string switches)
            {
                mFieldText = text;
                mFieldSwitches = switches;
            }

            ReplaceAction IReplacingCallback.Replacing(ReplacingArgs args)
            {
                // Create a builder to insert the field
                DocumentBuilder builder = new DocumentBuilder((Document) args.MatchNode.Document);
                // Move to the first node of the match
                builder.MoveTo(args.MatchNode);

                // If the user specified text to be used in the field as display text then use that, otherwise use the 
                // match string as the display text
                string insertText = !string.IsNullOrEmpty(mFieldText) ? mFieldText : args.Match.Value;

                // Insert the TC field before this node using the specified string as the display text and user defined switches
                builder.InsertField($"TC \"{insertText}\" {mFieldSwitches}");

                // We have done what we want so skip replacement
                return ReplaceAction.Skip;
            }
        }
        //ExEnd:InsertTCFieldHandler


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

        [Test]
        public static void CreateHeaderFooterUsingDocBuilder()
        {
            //ExStart:CreateHeaderFooterUsingDocBuilder
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            Section currentSection = builder.CurrentSection;
            PageSetup pageSetup = currentSection.PageSetup;

            // Specify if we want headers/footers of the first page to be different from other pages
            // You can also use PageSetup.OddAndEvenPagesHeaderFooter property to specify
            // Different headers/footers for odd and even pages
            pageSetup.DifferentFirstPageHeaderFooter = true;

            // --- Create header for the first page ---
            pageSetup.HeaderDistance = 20;
            builder.MoveToHeaderFooter(HeaderFooterType.HeaderFirst);
            builder.ParagraphFormat.Alignment = ParagraphAlignment.Center;

            // Set font properties for header text
            builder.Font.Name = "Arial";
            builder.Font.Bold = true;
            builder.Font.Size = 14;
            // Specify header title for the first page
            builder.Write("Aspose.Words Header/Footer Creation Primer - Title Page.");

            // --- Create header for pages other than first ---
            pageSetup.HeaderDistance = 20;
            builder.MoveToHeaderFooter(HeaderFooterType.HeaderPrimary);

            // Insert absolutely positioned image into the top/left corner of the header
            // Distance from the top/left edges of the page is set to 10 points
            string imageFileName = DocumentDir + "Aspose.Words.gif";
            builder.InsertImage(imageFileName, RelativeHorizontalPosition.Page, 10, RelativeVerticalPosition.Page, 10,
                50, 50, WrapType.Through);

            builder.ParagraphFormat.Alignment = ParagraphAlignment.Right;
            // Specify another header title for other pages
            builder.Write("Aspose.Words Header/Footer Creation Primer.");

            // --- Create footer for pages other than first ---
            builder.MoveToHeaderFooter(HeaderFooterType.FooterPrimary);

            // We use table with two cells to make one part of the text on the line (with page numbering)
            // To be aligned left, and the other part of the text (with copyright) to be aligned right
            builder.StartTable();

            // Clear table borders
            builder.CellFormat.ClearFormatting();

            builder.InsertCell();

            // Set first cell to 1/3 of the page width
            builder.CellFormat.PreferredWidth = PreferredWidth.FromPercent(100 / 3);

            // Insert page numbering text here
            // It uses PAGE and NUMPAGES fields to auto calculate current page number and total number of pages
            builder.Write("Page ");
            builder.InsertField("PAGE", "");
            builder.Write(" of ");
            builder.InsertField("NUMPAGES", "");

            // Align this text to the left
            builder.CurrentParagraph.ParagraphFormat.Alignment = ParagraphAlignment.Left;

            builder.InsertCell();
            // Set the second cell to 2/3 of the page width
            builder.CellFormat.PreferredWidth = PreferredWidth.FromPercent(100 * 2 / 3);

            builder.Write("(C) 2001 Aspose Pty Ltd. All rights reserved.");

            // Align this text to the right
            builder.CurrentParagraph.ParagraphFormat.Alignment = ParagraphAlignment.Right;

            builder.EndRow();
            builder.EndTable();

            builder.MoveToDocumentEnd();
            // Make page break to create a second page on which the primary headers/footers will be seen
            builder.InsertBreak(BreakType.PageBreak);

            // Make section break to create a third page with different page orientation
            builder.InsertBreak(BreakType.SectionBreakNewPage);

            // Get the new section and its page setup
            currentSection = builder.CurrentSection;
            pageSetup = currentSection.PageSetup;

            // Set page orientation of the new section to landscape
            pageSetup.Orientation = Orientation.Landscape;

            // This section does not need different first page header/footer
            // we need only one title page in the document and the header/footer for this page
            // has already been defined in the previous section
            pageSetup.DifferentFirstPageHeaderFooter = false;

            // This section displays headers/footers from the previous section by default
            // call currentSection.HeadersFooters.LinkToPrevious(false) to cancel this
            // page width is different for the new section and therefore we need to set 
            // a different cell widths for a footer table
            currentSection.HeadersFooters.LinkToPrevious(false);

            // If we want to use the already existing header/footer set for this section 
            // But with some minor modifications then it may be expedient to copy headers/footers
            // from the previous section and apply the necessary modifications where we want them
            CopyHeadersFootersFromPreviousSection(currentSection);

            // Find the footer that we want to change
            HeaderFooter primaryFooter = currentSection.HeadersFooters[HeaderFooterType.FooterPrimary];

            Row row = primaryFooter.Tables[0].FirstRow;
            row.FirstCell.CellFormat.PreferredWidth = PreferredWidth.FromPercent(100 / 3);
            row.LastCell.CellFormat.PreferredWidth = PreferredWidth.FromPercent(100 * 2 / 3);

            doc.Save(ArtifactsDir + "HeaderFooter.Primer.doc");
            //ExEnd:CreateHeaderFooterUsingDocBuilder
        }

        //ExStart:CopyHeadersFootersFromPreviousSection
        /// <summary>
        /// Clones and copies headers/footers form the previous section to the specified section.
        /// </summary>
        private static void CopyHeadersFootersFromPreviousSection(Section section)
        {
            Section previousSection = (Section) section.PreviousSibling;

            if (previousSection == null)
                return;

            section.HeadersFooters.Clear();

            foreach (HeaderFooter headerFooter in previousSection.HeadersFooters)
                section.HeadersFooters.Add(headerFooter.Clone(true));
        }
        //ExEnd:CopyHeadersFootersFromPreviousSection
    }
}