using Aspose.Words.Saving;
using NUnit.Framework;

namespace Aspose.Words.Examples.CSharp
{
    class EmbeddingWindowsStandardFonts : TestDataHelper
    {
        [Test]
        public static void AvoidEmbeddingCoreFonts()
        {
            //ExStart:AvoidEmbeddingCoreFonts
            Document doc = new Document(RenderingPrintingDir + "Rendering.doc");

            // To disable embedding of core fonts and substitute PDF type 1 fonts set UseCoreFonts to true
            PdfSaveOptions options = new PdfSaveOptions();
            options.UseCoreFonts = true;

            // The output PDF will not be embedded with core fonts such as Arial, Times New Roman etc
            doc.Save(ArtifactsDir + "AvoidEmbeddingCoreFonts.pdf", options);
            //ExEnd:AvoidEmbeddingCoreFonts
        }

        [Test]
        public static void SkipEmbeddedArialAndTimesRomanFonts()
        {
            //ExStart:SkipEmbeddedArialAndTimesRomanFonts
            Document doc = new Document(RenderingPrintingDir + "Rendering.doc");
            
            // To subset fonts in the output PDF document, simply create new PdfSaveOptions and set EmbedFullFonts to false
            // To disable embedding standard windows font use the PdfSaveOptions and set the EmbedStandardWindowsFonts property to false
            PdfSaveOptions options = new PdfSaveOptions();
            options.FontEmbeddingMode = PdfFontEmbeddingMode.EmbedAll;

            // The output PDF will be saved without embedding standard windows fonts
            doc.Save(ArtifactsDir + "SkipEmbeddedArialAndTimesRomanFonts.pdf");
            //ExEnd:SkipEmbeddedArialAndTimesRomanFonts
        }
    }
}