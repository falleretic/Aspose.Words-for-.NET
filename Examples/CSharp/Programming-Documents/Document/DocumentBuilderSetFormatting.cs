using Aspose.Words.Tables;
using NUnit.Framework;

namespace Aspose.Words.Examples.CSharp.Programming_Documents.Working_With_Document
{
    class DocumentBuilderSetFormatting : TestDataHelper
    {
        [Test]
        public static void SetSpaceBetweenAsianAndLatinText()
        {
            //ExStart:DocumentBuilderSetSpacebetweenAsianandLatintext
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Set paragraph formatting properties
            ParagraphFormat paragraphFormat = builder.ParagraphFormat;
            paragraphFormat.AddSpaceBetweenFarEastAndAlpha = true;
            paragraphFormat.AddSpaceBetweenFarEastAndDigit = true;

            builder.Writeln("Automatically adjust space between Asian and Latin text");
            builder.Writeln("Automatically adjust space between Asian text and numbers");

            doc.Save(ArtifactsDir + "DocumentBuilderSetSpaceBetweenAsianAndLatinText.doc");
            //ExEnd:DocumentBuilderSetSpacebetweenAsianandLatintext
        }

        [Test]
        public static void SetAsianTypographyLinebreakGroupProp()
        {
            //ExStart:SetAsianTypographyLinebreakGroupProp
            Document doc = new Document(DocumentDir + "Input.docx");

            ParagraphFormat format = doc.FirstSection.Body.Paragraphs[0].ParagraphFormat;
            format.FarEastLineBreakControl = false;
            format.WordWrap = true;
            format.HangingPunctuation = false;

            doc.Save(ArtifactsDir + "SetAsianTypographyLinebreakGroupProp.docx");
            //ExEnd:SetAsianTypographyLinebreakGroupProp
        }

        [Test]
        public static void SetFontFormatting()
        {
            //ExStart:DocumentBuilderSetFontFormatting
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Set font formatting properties
            Font font = builder.Font;
            font.Bold = true;
            font.Color = System.Drawing.Color.DarkBlue;
            font.Italic = true;
            font.Name = "Arial";
            font.Size = 24;
            font.Spacing = 5;
            font.Underline = Underline.Double;

            // Output formatted text
            builder.Writeln("I'm a very nice formatted string.");
            
            doc.Save(ArtifactsDir + "DocumentBuilderSetFontFormatting.doc");
            //ExEnd:DocumentBuilderSetFontFormatting
        }

        [Test]
        public static void SetParagraphFormatting()
        {
            //ExStart:DocumentBuilderSetParagraphFormatting
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Set paragraph formatting properties
            ParagraphFormat paragraphFormat = builder.ParagraphFormat;
            paragraphFormat.Alignment = ParagraphAlignment.Center;
            paragraphFormat.LeftIndent = 50;
            paragraphFormat.RightIndent = 50;
            paragraphFormat.SpaceAfter = 25;

            // Output text
            builder.Writeln(
                "I'm a very nice formatted paragraph. I'm intended to demonstrate how the left and right indents affect word wrapping.");
            builder.Writeln(
                "I'm another nice formatted paragraph. I'm intended to demonstrate how the space after paragraph looks like.");

            doc.Save(ArtifactsDir + "DocumentBuilderSetParagraphFormatting.doc");
            //ExEnd:DocumentBuilderSetParagraphFormatting
        }

        [Test]
        public static void SetTableCellFormatting()
        {
            //ExStart:DocumentBuilderSetTableCellFormatting
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            builder.StartTable();
            builder.InsertCell();

            // Set the cell formatting
            CellFormat cellFormat = builder.CellFormat;
            cellFormat.Width = 250;
            cellFormat.LeftPadding = 30;
            cellFormat.RightPadding = 30;
            cellFormat.TopPadding = 30;
            cellFormat.BottomPadding = 30;

            builder.Writeln("I'm a wonderful formatted cell.");

            builder.EndRow();
            builder.EndTable();

            doc.Save(ArtifactsDir + "DocumentBuilderSetTableCellFormatting.doc");
            //ExEnd:DocumentBuilderSetTableCellFormatting
        }

        [Test]
        public static void SetTableRowFormatting()
        {
            //ExStart:DocumentBuilderSetTableRowFormatting
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            Table table = builder.StartTable();
            builder.InsertCell();

            // Set the row formatting
            RowFormat rowFormat = builder.RowFormat;
            rowFormat.Height = 100;
            rowFormat.HeightRule = HeightRule.Exactly;
            // These formatting properties are set on the table and are applied to all rows in the table
            table.LeftPadding = 30;
            table.RightPadding = 30;
            table.TopPadding = 30;
            table.BottomPadding = 30;

            builder.Writeln("I'm a wonderful formatted row.");

            builder.EndRow();
            builder.EndTable();

            doc.Save(ArtifactsDir + "DocumentBuilderSetTableRowFormatting.doc");
            //ExEnd:DocumentBuilderSetTableRowFormatting
        }

        [Test]
        public static void SetMultilevelListFormatting()
        {
            //ExStart:DocumentBuilderSetMultilevelListFormatting
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            builder.ListFormat.ApplyNumberDefault();
            builder.Writeln("Item 1");
            builder.Writeln("Item 2");

            builder.ListFormat.ListIndent();
            builder.Writeln("Item 2.1");
            builder.Writeln("Item 2.2");
            
            builder.ListFormat.ListIndent();
            builder.Writeln("Item 2.2.1");
            builder.Writeln("Item 2.2.2");

            builder.ListFormat.ListOutdent();
            builder.Writeln("Item 2.3");

            builder.ListFormat.ListOutdent();
            builder.Writeln("Item 3");

            builder.ListFormat.RemoveNumbers();
            
            doc.Save(ArtifactsDir + "DocumentBuilderSetMultilevelListFormatting.doc");
            //ExEnd:DocumentBuilderSetMultilevelListFormatting
        }

        [Test]
        public static void SetPageSetupAndSectionFormatting()
        {
            //ExStart:DocumentBuilderSetPageSetupAndSectionFormatting
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Set page properties
            builder.PageSetup.Orientation = Orientation.Landscape;
            builder.PageSetup.LeftMargin = 50;
            builder.PageSetup.PaperSize = PaperSize.Paper10x14;

            doc.Save(ArtifactsDir + "DocumentBuilderSetPageSetupAndSectionFormatting.doc");
            //ExEnd:DocumentBuilderSetPageSetupAndSectionFormatting
        }

        [Test]
        public static void ApplyParagraphStyle()
        {
            //ExStart:DocumentBuilderApplyParagraphStyle
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Set paragraph style
            builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Title;
            builder.Write("Hello");
            
            doc.Save(ArtifactsDir + "DocumentBuilderApplyParagraphStyle.doc");
            //ExEnd:DocumentBuilderApplyParagraphStyle
        }

        [Test]
        public static void ApplyBordersAndShadingToParagraph()
        {
            //ExStart:DocumentBuilderApplyBordersAndShadingToParagraph
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Set paragraph borders
            BorderCollection borders = builder.ParagraphFormat.Borders;
            borders.DistanceFromText = 20;
            borders[BorderType.Left].LineStyle = LineStyle.Double;
            borders[BorderType.Right].LineStyle = LineStyle.Double;
            borders[BorderType.Top].LineStyle = LineStyle.Double;
            borders[BorderType.Bottom].LineStyle = LineStyle.Double;

            // Set paragraph shading
            Shading shading = builder.ParagraphFormat.Shading;
            shading.Texture = TextureIndex.TextureDiagonalCross;
            shading.BackgroundPatternColor = System.Drawing.Color.LightCoral;
            shading.ForegroundPatternColor = System.Drawing.Color.LightSalmon;

            builder.Write("I'm a formatted paragraph with double border and nice shading.");
            
            doc.Save(ArtifactsDir + "DocumentBuilderApplyBordersAndShadingToParagraph.doc");
            //ExEnd:DocumentBuilderApplyBordersAndShadingToParagraph
        }
        
        [Test]
        public static void ChangeAsianParagraphSpacingandIndents()
        {
            // ExStart:ChangeAsianParagraphSpacingandIndents
            Document doc = new Document(DocumentDir + "Input.docx");

            ParagraphFormat format = doc.FirstSection.Body.FirstParagraph.ParagraphFormat;
            format.CharacterUnitLeftIndent = 10;       // ParagraphFormat.LeftIndent will be updated
            format.CharacterUnitRightIndent = 10;      // ParagraphFormat.RightIndent will be updated
            format.CharacterUnitFirstLineIndent = 20;  // ParagraphFormat.FirstLineIndent will be updated
            format.LineUnitBefore = 5;                 // ParagraphFormat.SpaceBefore will be updated
            format.LineUnitAfter = 10;                 // ParagraphFormat.SpaceAfter will be updated

            doc.Save(ArtifactsDir + "ChangeAsianParagraphSpacingandIndents.doc");
            // ExEnd:ChangeAsianParagraphSpacingandIndents
        }

        [Test]
        public static void SetSnapToGrid()
        {
            // ExStart:SetSnapToGrid
            Document doc = new Document();

            Paragraph par = doc.FirstSection.Body.FirstParagraph;
            par.ParagraphFormat.SnapToGrid = true;
            par.Runs[0].Font.SnapToGrid = true;

            doc.Save(ArtifactsDir + "SetSnapToGrid.doc");
            // ExEnd:SetSnapToGrid
        }

        [Test]
        public static void SetFontEmphasisMark()
        {
            // ExStart:SetFontEmphasisMark
            Document document = new Document();
            DocumentBuilder builder = new DocumentBuilder(document);

            builder.Font.EmphasisMark = EmphasisMark.UnderSolidCircle;

            builder.Write("Emphasis text");
            builder.Writeln();
            builder.Font.ClearFormatting();
            builder.Write("Simple text");

            document.Save(ArtifactsDir + "FontEmphasisMark.doc");
            // ExEnd:SetFontEmphasisMark
        }
    }
}