using System.Drawing;
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
    }
}
