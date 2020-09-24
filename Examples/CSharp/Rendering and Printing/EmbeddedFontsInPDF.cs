using Aspose.Words.Saving;
using NUnit.Framework;

namespace Aspose.Words.Examples.CSharp
{
    class EmbeddedFontsInPdf : TestDataHelper
    {
        [Test]
        public static void EmbeddedAllFonts()
        {
            //ExStart:EmbeddAllFonts
            Document doc = new Document(RenderingPrintingDir + "Rendering.docx");

            // Aspose.Words embeds full fonts by default when EmbedFullFonts is set to true. The property below can be changed
            // Each time a document is rendered
            PdfSaveOptions options = new PdfSaveOptions();
            options.EmbedFullFonts = true;

            // The output PDF will be embedded with all fonts found in the document
            doc.Save(ArtifactsDir + "EmbeddedFontsInPdf.pdf", options);
            //ExEnd:EmbeddAllFonts
        }

        [Test]
        public static void EmbeddedSubsetFonts()
        {
            //ExStart:EmbeddSubsetFonts
            Document doc = new Document(RenderingPrintingDir + "Rendering.docx");
            
            // To subset fonts in the output PDF document, simply create new PdfSaveOptions and set EmbedFullFonts to false
            PdfSaveOptions options = new PdfSaveOptions();
            options.EmbedFullFonts = false;
            
            // The output PDF will contain subsets of the fonts in the document. Only the glyphs used
            // in the document are included in the PDF fonts
            doc.Save(ArtifactsDir + "EmbeddSubsetFonts.pdf", options);
            //ExEnd:EmbeddSubsetFonts
        }

        [Test]
        public static void SetFontEmbeddingMode()
        {
            // ExStart:SetFontEmbeddingMode
            // Load the document to render.
            Document doc = new Document(RenderingPrintingDir + "Rendering.docx");

            // To disable embedding standard windows font use the PdfSaveOptions and set the EmbedStandardWindowsFonts property to false.
            PdfSaveOptions options = new PdfSaveOptions();
            options.FontEmbeddingMode = PdfFontEmbeddingMode.EmbedNone;

            // The output PDF will be saved without embedding standard windows fonts.
            doc.Save(ArtifactsDir + "Rendering.DisableEmbedWindowsFonts.pdf");
            // ExEnd:SetFontEmbeddingMode
        }
    }
}