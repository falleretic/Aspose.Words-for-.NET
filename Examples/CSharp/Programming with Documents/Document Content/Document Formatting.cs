using System;
using Aspose.Words.Tables;
using NUnit.Framework;

namespace Aspose.Words.Examples.CSharp.DocumentEx
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
            Document doc = new Document(DocumentDir + "Asian typography.docx");

            ParagraphFormat format = doc.FirstSection.Body.Paragraphs[0].ParagraphFormat;
            format.FarEastLineBreakControl = false;
            format.WordWrap = true;
            format.HangingPunctuation = false;

            doc.Save(ArtifactsDir + "SetAsianTypographyLinebreakGroupProp.docx");
            //ExEnd:SetAsianTypographyLinebreakGroupProp
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
            Document doc = new Document(DocumentDir + "Asian typography.docx");

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
        public static void ParagraphStyleSeparator()
        {
            //ExStart:ParagraphStyleSeparator
            Document doc = new Document(DocumentDir + "Document.docx");

            foreach (Paragraph paragraph in doc.GetChildNodes(NodeType.Paragraph, true))
            {
                if (paragraph.BreakIsStyleSeparator)
                {
                    Console.WriteLine("Separator Found!");
                }
            }
            //ExEnd:ParagraphStyleSeparator
        }
    }
}