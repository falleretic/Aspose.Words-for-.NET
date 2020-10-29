using System;
using System.Drawing;
using System.Text.RegularExpressions;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Words.Fields;
using Aspose.Words.Replacing;
using Aspose.Words.Tables;
using NUnit.Framework;
using Font = Aspose.Words.Font;

namespace SiteExamples.Programming_with_Documents.Document_Content
{
    class AddContentUsingDocumentBuilder : SiteExamplesBase
    {
        [Test]
        public static void BuildTable()
        {
            //ExStart:BuildTable
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            Table table = builder.StartTable();
            builder.InsertCell();
            table.AutoFit(AutoFitBehavior.FixedColumnWidths);

            builder.CellFormat.VerticalAlignment = CellVerticalAlignment.Center;
            builder.Write("This is row 1 cell 1");

            builder.InsertCell();
            builder.Write("This is row 1 cell 2");

            builder.EndRow();

            builder.InsertCell();
            
            builder.RowFormat.Height = 100;
            builder.RowFormat.HeightRule = HeightRule.Exactly;
            builder.CellFormat.Orientation = TextOrientation.Upward;
            builder.Writeln("This is row 2 cell 1");

            builder.InsertCell();
            builder.CellFormat.Orientation = TextOrientation.Downward;
            builder.Writeln("This is row 2 cell 2");

            builder.EndRow();
            builder.EndTable();

            doc.Save(ArtifactsDir + "DocumentBuilder.BuildTable.docx");
            //ExEnd:BuildTable
        }

        [Test]
        public static void InsertHorizontalRule()
        {
            //ExStart:InsertHorizontalRule
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            builder.Writeln("Insert a horizontal rule shape into the document.");
            builder.InsertHorizontalRule();

            doc.Save(ArtifactsDir + "DocumentBuilder.InsertHorizontalRule.docx");
            //ExEnd:InsertHorizontalRule
        }

        [Test]
        public static void HorizontalRuleFormat()
        {
            //ExStart:HorizontalRuleFormat
            DocumentBuilder builder = new DocumentBuilder();

            Shape shape = builder.InsertHorizontalRule();
            
            HorizontalRuleFormat horizontalRuleFormat = shape.HorizontalRuleFormat;
            horizontalRuleFormat.Alignment = HorizontalRuleAlignment.Center;
            horizontalRuleFormat.WidthPercent = 70;
            horizontalRuleFormat.Height = 3;
            horizontalRuleFormat.Color = Color.Blue;
            horizontalRuleFormat.NoShade = true;

            builder.Document.Save(ArtifactsDir + "DocumentBuilder.HorizontalRuleFormat.docx");
            //ExEnd:HorizontalRuleFormat
        }

        [Test]
        public static void InsertBreak()
        {
            //ExStart:InsertBreak
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            builder.Writeln("This is page 1.");
            builder.InsertBreak(BreakType.PageBreak);

            builder.Writeln("This is page 2.");
            builder.InsertBreak(BreakType.PageBreak);

            builder.Writeln("This is page 3.");
            doc.Save(ArtifactsDir + "DocumentBuilder.InsertBreak.docx");
            //ExEnd:InsertBreak
        }

        [Test]
        public static void InsertTextInputFormField()
        {
            //ExStart:InsertTextInputFormField
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);
            
            builder.InsertTextInput("TextInput", TextFormFieldType.Regular, "", "Hello", 0);

            doc.Save(ArtifactsDir + "DocumentBuilder.InsertTextInputFormField.docx");
            //ExEnd:InsertTextInputFormField
        }

        [Test]
        public static void InsertCheckBoxFormField()
        {
            //ExStart:InsertCheckBoxFormField
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);
            
            builder.InsertCheckBox("CheckBox", true, true, 0);

            doc.Save(ArtifactsDir + "DocumentBuilder.InsertCheckBoxFormField.docx");
            //ExEnd:InsertCheckBoxFormField
        }

        [Test]
        public static void InsertComboBoxFormField()
        {
            //ExStart:InsertComboBoxFormField
            string[] items = { "One", "Two", "Three" };

            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            builder.InsertComboBox("DropDown", items, 0);

            doc.Save(ArtifactsDir + "DocumentBuilder.InsertComboBoxFormField.docx");
            //ExEnd:InsertComboBoxFormField
        }

        [Test]
        public static void InsertHtml()
        {
            //ExStart:InsertHtml
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);
            builder.InsertHtml(
                "<P align='right'>Paragraph right</P>" +
                "<b>Implicit paragraph left</b>" +
                "<div align='center'>Div center</div>" +
                "<h1 align='left'>Heading 1 left.</h1>");

            doc.Save(ArtifactsDir + "DocumentBuilder.InsertHtml.docx");
            //ExEnd:InsertHtml
        }

        [Test]
        public static void InsertHyperlink()
        {
            //ExStart:InsertHyperlink
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);
            
            builder.Write("Please make sure to visit ");
            builder.Font.Color = Color.Blue;
            builder.Font.Underline = Underline.Single;
            
            builder.InsertHyperlink("Aspose Website", "http://www.aspose.com", false);
            
            builder.Font.ClearFormatting();
            builder.Write(" for more information.");

            doc.Save(ArtifactsDir + "DocumentBuilder.InsertHyperlink.docx");
            //ExEnd:InsertHyperlink
        }

        [Test]
        public static void InsertTableOfContents()
        {
            //ExStart:InsertTableOfContents
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);
            
            builder.InsertTableOfContents("\\o \"1-3\" \\h \\z \\u");
            // Start the actual document content on the second page.
            builder.InsertBreak(BreakType.PageBreak);

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

            //ExStart:UpdateFields
            // The newly inserted table of contents will be initially empty.
            // It needs to be populated by updating the fields in the document.
            doc.UpdateFields();
            //ExEnd:UpdateFields

            doc.Save(ArtifactsDir + "DocumentBuilder.InsertTableOfContents.docx");
            //ExEnd:InsertTableOfContents
        }

        [Test]
        public static void InsertInlineImage()
        {
            //ExStart:InsertInlineImage
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            builder.InsertImage(ImagesDir + "Transparent background logo.png");

            doc.Save(ArtifactsDir + "DocumentBuilder.InsertInlineImage.docx");
            //ExEnd:InsertInlineImage
        }

        [Test]
        public static void InsertFloatingImage()
        {
            //ExStart:InsertFloatingImage
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            builder.InsertImage(ImagesDir + "Transparent background logo.png",
                RelativeHorizontalPosition.Margin,
                100,
                RelativeVerticalPosition.Margin,
                100,
                200,
                100,
                WrapType.Square);

            doc.Save(ArtifactsDir + "DocumentBuilder.InsertFloatingImage.docx");
            //ExEnd:InsertFloatingImage
        }

        [Test]
        public static void InsertParagraph()
        {
            //ExStart:InsertParagraph
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            Font font = builder.Font;
            font.Size = 16;
            font.Bold = true;
            font.Color = Color.Blue;
            font.Name = "Arial";
            font.Underline = Underline.Dash;

            ParagraphFormat paragraphFormat = builder.ParagraphFormat;
            paragraphFormat.FirstLineIndent = 8;
            paragraphFormat.Alignment = ParagraphAlignment.Justify;
            paragraphFormat.KeepTogether = true;

            builder.Writeln("A whole paragraph.");

            doc.Save(ArtifactsDir + "DocumentBuilder.InsertParagraph.docx");
            //ExEnd:InsertParagraph
        }

        [Test]
        public static void InsertTCField()
        {
            //ExStart:InsertTCField
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            builder.InsertField("TC \"Entry Text\" \\f t");

            doc.Save(ArtifactsDir + "DocumentBuilder.InsertTCField.docx");
            //ExEnd:InsertTCField
        }

        [Test]
        public static void InsertTCFieldsAtText()
        {
            //ExStart:InsertTCFieldsAtText
            Document doc = new Document();

            FindReplaceOptions options = new FindReplaceOptions();
            options.ApplyFont.HighlightColor = Color.DarkOrange;
            options.ReplacingCallback = new InsertTCFieldHandler("Chapter 1", "\\l 1");

            doc.Range.Replace(new Regex("The Beginning"), "", options);
            //ExEnd:InsertTCFieldsAtText
        }

        //ExStart:InsertTCFieldHandler
        public sealed class InsertTCFieldHandler : IReplacingCallback
        {
            // Store the text and switches to be used for the TC fields.
            private readonly string mFieldText;
            private readonly string mFieldSwitches;

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
                DocumentBuilder builder = new DocumentBuilder((Document) args.MatchNode.Document);
                builder.MoveTo(args.MatchNode);

                // If the user specified text to be used in the field as display text then use that, otherwise use the 
                // match string as the display text.
                string insertText = !string.IsNullOrEmpty(mFieldText) ? mFieldText : args.Match.Value;

                builder.InsertField($"TC \"{insertText}\" {mFieldSwitches}");

                return ReplaceAction.Skip;
            }
        }
        //ExEnd:InsertTCFieldHandler
        
        [Test]
        public static void CursorPosition()
        {
            //ExStart:CursorPosition
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            Node curNode = builder.CurrentNode;
            Paragraph curParagraph = builder.CurrentParagraph;
            //ExEnd:CursorPosition

            Console.WriteLine("\nCursor move to paragraph: " + curParagraph.GetText());
        }

        [Test]
        public static void MoveToNode()
        {
            //ExStart:MoveToNode
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            builder.MoveTo(doc.FirstSection.Body.LastParagraph);
            // ExEnd:MoveToNode
        }

        [Test]
        public static void MoveToDocumentStartEnd()
        {
            //ExStart:MoveToDocumentStartEnd
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            builder.MoveToDocumentEnd();
            Console.WriteLine("\nThis is the end of the document.");

            builder.MoveToDocumentStart();
            Console.WriteLine("\nThis is the beginning of the document.");
            //ExEnd:MoveToDocumentStartEnd            
        }

        [Test]
        public static void MoveToSection()
        {
            //ExStart:MoveToSection
            Document doc = new Document();
            doc.AppendChild(new Section(doc));

            DocumentBuilder builder = new DocumentBuilder(doc);

            // Parameters are 0-index. Moves to third section.
            builder.MoveToSection(1);
            builder.Writeln("This is the 2rd section.");
            //ExEnd:MoveToSection               
        }

        [Test]
        public static void HeadersAndFooters()
        {
            //ExStart:HeadersAndFooters
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            builder.PageSetup.DifferentFirstPageHeaderFooter = true;
            builder.PageSetup.OddAndEvenPagesHeaderFooter = true;

            builder.MoveToHeaderFooter(HeaderFooterType.HeaderFirst);
            builder.Write("Header First");
            builder.MoveToHeaderFooter(HeaderFooterType.HeaderEven);
            builder.Write("Header Even");
            builder.MoveToHeaderFooter(HeaderFooterType.HeaderPrimary);
            builder.Write("Header Odd");

            builder.MoveToSection(0);
            builder.Writeln("Page1");
            builder.InsertBreak(BreakType.PageBreak);
            builder.Writeln("Page2");
            builder.InsertBreak(BreakType.PageBreak);
            builder.Writeln("Page3");

            doc.Save(ArtifactsDir + "DocumentBuilder.HeadersAndFooters.doc");
            //ExEnd:HeadersAndFooters
        }

        [Test]
        public static void MoveToParagraph()
        {
            //ExStart:MoveToParagraph
            Document doc = new Document(MyDir + "Paragraphs.docx");
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Parameters are 0-index. Moves to third paragraph.
            builder.MoveToParagraph(2, 0);
            builder.Writeln("This is the 3rd paragraph.");
            //ExEnd:MoveToParagraph               
        }

        [Test]
        public static void MoveToTableCell()
        {
            //ExStart:MoveToTableCell
            Document doc = new Document(MyDir + "Tables.docx");
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Move the builder to row 3, cell 4 of the first table.
            builder.MoveToCell(0, 2, 3, 0);
            builder.Writeln("Hello World!");
            //ExEnd:MoveToTableCell               
        }

        [Test]
        public static void MoveToBookmark()
        {
            //ExStart:MoveToBookmark
            Document doc = new Document(MyDir + "Bookmarks.docx");
            DocumentBuilder builder = new DocumentBuilder(doc);

            builder.MoveToBookmark("MyBookmark1");
            builder.Writeln("This is a very cool bookmark.");
            //ExEnd:MoveToBookmark               
        }

        [Test]
        public static void MoveToBookmarkEnd()
        {
            //ExStart:MoveToBookmarkEnd
            Document doc = new Document(MyDir + "Bookmarks.docx");
            DocumentBuilder builder = new DocumentBuilder(doc);

            builder.MoveToBookmark("MyBookmark1", false, true);
            builder.Writeln("This is a very cool bookmark.");
            //ExEnd:MoveToBookmarkEnd              
        }

        [Test]
        public static void MoveToMergeField()
        {
            //ExStart:MoveToMergeField
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            builder.InsertField(@"MERGEFIELD MyMergeField1 \* MERGEFORMAT");
            builder.InsertField(@"MERGEFIELD MyMergeField2 \* MERGEFORMAT");

            builder.MoveToMergeField("MyMergeField1");
            builder.Writeln("This is a very nice merge field.");
            //ExEnd:MoveToMergeField              
        }

        [Test]
        public static void CreateHeaderFooterUsingDocBuilder()
        {
            //ExStart:CreateHeaderFooterUsingDocBuilder
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            Section currentSection = builder.CurrentSection;
            PageSetup pageSetup = currentSection.PageSetup;
            // Specify if we want headers/footers of the first page to be different from other pages.
            // You can also use PageSetup.OddAndEvenPagesHeaderFooter property to specify.
            // Different headers/footers for odd and even pages.
            pageSetup.DifferentFirstPageHeaderFooter = true;
            pageSetup.HeaderDistance = 20;

            builder.MoveToHeaderFooter(HeaderFooterType.HeaderFirst);
            builder.ParagraphFormat.Alignment = ParagraphAlignment.Center;

            builder.Font.Name = "Arial";
            builder.Font.Bold = true;
            builder.Font.Size = 14;
            
            builder.Write("Aspose.Words Header/Footer Creation Primer - Title Page.");

            pageSetup.HeaderDistance = 20;
            builder.MoveToHeaderFooter(HeaderFooterType.HeaderPrimary);

            // Insert absolutely positioned image into the top/left corner of the header.
            // Distance from the top/left edges of the page is set to 10 points.
            builder.InsertImage(ImagesDir + "Graphics Interchange Format.gif", RelativeHorizontalPosition.Page, 10, 
                RelativeVerticalPosition.Page, 10, 50, 50, WrapType.Through);

            builder.ParagraphFormat.Alignment = ParagraphAlignment.Right;
            
            builder.Write("Aspose.Words Header/Footer Creation Primer.");

            builder.MoveToHeaderFooter(HeaderFooterType.FooterPrimary);

            // We use table with two cells to make one part of the text on the line (with page numbering).
            // To be aligned left, and the other part of the text (with copyright) to be aligned right.
            builder.StartTable();

            builder.CellFormat.ClearFormatting();

            builder.InsertCell();

            builder.CellFormat.PreferredWidth = PreferredWidth.FromPercent(100 / 3);

            // It uses PAGE and NUMPAGES fields to auto calculate current page number and total number of pages.
            builder.Write("Page ");
            builder.InsertField("PAGE", "");
            builder.Write(" of ");
            builder.InsertField("NUMPAGES", "");

            builder.CurrentParagraph.ParagraphFormat.Alignment = ParagraphAlignment.Left;

            builder.InsertCell();
            
            builder.CellFormat.PreferredWidth = PreferredWidth.FromPercent(100 * 2 / 3);

            builder.Write("(C) 2001 Aspose Pty Ltd. All rights reserved.");

            builder.CurrentParagraph.ParagraphFormat.Alignment = ParagraphAlignment.Right;

            builder.EndRow();
            builder.EndTable();

            builder.MoveToDocumentEnd();
            
            // Make page break to create a second page on which the primary headers/footers will be seen.
            builder.InsertBreak(BreakType.PageBreak);
            builder.InsertBreak(BreakType.SectionBreakNewPage);

            currentSection = builder.CurrentSection;
            pageSetup = currentSection.PageSetup;
            pageSetup.Orientation = Orientation.Landscape;
            // This section does not need different first page header/footer
            // we need only one title page in the document and the header/footer for this page
            // has already been defined in the previous section.
            pageSetup.DifferentFirstPageHeaderFooter = false;

            // This section displays headers/footers from the previous section by default
            // call currentSection.HeadersFooters.LinkToPrevious(false) to cancel this
            // page width is different for the new section and therefore we need to set 
            // a different cell widths for a footer table.
            currentSection.HeadersFooters.LinkToPrevious(false);

            // If we want to use the already existing header/footer set for this section.
            // But with some minor modifications then it may be expedient to copy headers/footers
            // from the previous section and apply the necessary modifications where we want them.
            CopyHeadersFootersFromPreviousSection(currentSection);

            HeaderFooter primaryFooter = currentSection.HeadersFooters[HeaderFooterType.FooterPrimary];

            Row row = primaryFooter.Tables[0].FirstRow;
            row.FirstCell.CellFormat.PreferredWidth = PreferredWidth.FromPercent(100 / 3);
            row.LastCell.CellFormat.PreferredWidth = PreferredWidth.FromPercent(100 * 2 / 3);

            doc.Save(ArtifactsDir + "DocumentBuilder.CreateHeaderFooterUsingDocBuilder.docx");
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