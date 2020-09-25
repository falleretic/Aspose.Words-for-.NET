using System.Drawing;
using Aspose.Words.Drawing;
using NUnit.Framework;

namespace Aspose.Words.Examples.CSharp.DocumentEx
{
    class WorkWithWatermark : TestDataHelper
    {
        [Test]
        public static void AddTextWatermarkWithSpecificOptions()
        {
            //ExStart:AddTextWatermarkWithSpecificOptions
            Document doc = new Document(DocumentDir + "Document.docx");

            TextWatermarkOptions options = new TextWatermarkOptions()
            {
                FontFamily = "Arial",
                FontSize = 36,
                Color = Color.Black,
                Layout = WatermarkLayout.Horizontal,
                IsSemitrasparent = false
            };

            doc.Watermark.SetText("Test", options);

            doc.Save(ArtifactsDir + "AddTextWatermark.docx");
            //ExEnd:AddTextWatermarkWithSpecificOptions
        }

        [Test]
        public static void AddImageWatermarkWithSpecificOptions()
        {
            //ExStart:AddImageWatermarkWithSpecificOptions
            Document doc = new Document(DocumentDir + "Document.docx");

            ImageWatermarkOptions options = new ImageWatermarkOptions()
            {
                Scale = 5,
                IsWashout = false
            };

            doc.Watermark.SetImage(Image.FromFile(DocumentDir + "Watermark.png"), options);

            doc.Save(ArtifactsDir + "AddImageWatermark.docx");
            //ExEnd:AddImageWatermarkWithSpecificOptions
        }

        [Test]
        public static void RemoveWatermarkFromDocument()
        {
            //ExStart:RemoveWatermarkFromDocument
            Document doc = new Document(ArtifactsDir + "AddTextWatermark.docx");

            if (doc.Watermark.Type == WatermarkType.Text)
            {
                doc.Watermark.Remove();
            }

            doc.Save(ArtifactsDir + "RemoveWatermark_out.docx");
            //ExEnd:RemoveWatermarkFromDocument
        }

        //ExStart:AddWatermark
        [Test]
        public static void AddWatermark()
        {
            Document doc = new Document(ImagesDir + "Document.docx");

            InsertWatermarkText(doc, "CONFIDENTIAL");

            doc.Save(ArtifactsDir + "TestFile.Watermark.docx");
        }

        /// <summary>
        /// Inserts a watermark into a document.
        /// </summary>
        /// <param name="doc">The input document.</param>
        /// <param name="watermarkText">Text of the watermark.</param>
        private static void InsertWatermarkText(Document doc, string watermarkText)
        {
            // Create a watermark shape, this will be a WordArt shape
            // You are free to try other shape types as watermarks
            Shape watermark = new Shape(doc, ShapeType.TextPlainText);
            watermark.Name = "WaterMark";

            // Set up the text of the watermark
            watermark.TextPath.Text = watermarkText;
            watermark.TextPath.FontFamily = "Arial";
            watermark.Width = 500;
            watermark.Height = 100;

            // Text will be directed from the bottom-left to the top-right corner
            watermark.Rotation = -40;

            // Remove the following two lines if you need a solid black text
            watermark.Fill.Color = Color.Gray; // Try LightGray to get more Word-style watermark
            watermark.StrokeColor = Color.Gray; // Try LightGray to get more Word-style watermark

            // Place the watermark in the page center
            watermark.RelativeHorizontalPosition = RelativeHorizontalPosition.Page;
            watermark.RelativeVerticalPosition = RelativeVerticalPosition.Page;
            watermark.WrapType = WrapType.None;
            watermark.VerticalAlignment = VerticalAlignment.Center;
            watermark.HorizontalAlignment = HorizontalAlignment.Center;

            // Create a new paragraph and append the watermark to this paragraph
            Paragraph watermarkPara = new Paragraph(doc);
            watermarkPara.AppendChild(watermark);

            // Insert the watermark into all headers of each document section
            foreach (Section sect in doc.Sections)
            {
                // There could be up to three different headers in each section, since we want
                // The watermark to appear on all pages, insert into all headers
                InsertWatermarkIntoHeader(watermarkPara, sect, HeaderFooterType.HeaderPrimary);
                InsertWatermarkIntoHeader(watermarkPara, sect, HeaderFooterType.HeaderFirst);
                InsertWatermarkIntoHeader(watermarkPara, sect, HeaderFooterType.HeaderEven);
            }
        }

        private static void InsertWatermarkIntoHeader(Paragraph watermarkPara, Section sect,
            HeaderFooterType headerType)
        {
            HeaderFooter header = sect.HeadersFooters[headerType];

            if (header == null)
            {
                // There is no header of the specified type in the current section, create it
                header = new HeaderFooter(sect.Document, headerType);
                sect.HeadersFooters.Add(header);
            }

            // Insert a clone of the watermark into the header
            header.AppendChild(watermarkPara.Clone(true));
        }
        //ExEnd:AddWatermark

        [Test]
        //ExStart:RemoveWatermark
        public static void RemoveWatermark()
        {
            Document doc = new Document(ArtifactsDir + "TestFile.Watermark.docx");
            
            RemoveWatermarkText(doc);
            
            doc.Save(ArtifactsDir + "RemoveWatermark.docx");
        }

        private static void RemoveWatermarkText(Document doc)
        {
            foreach (HeaderFooter hf in doc.GetChildNodes(NodeType.HeaderFooter, true))
            {
                foreach (Shape shape in hf.GetChildNodes(NodeType.Shape, true))
                {
                    if (shape.Name.Contains("WaterMark"))
                    {
                        shape.Remove();
                    }
                }
            }
        }
        //ExEnd:RemoveWatermark
    }
}
