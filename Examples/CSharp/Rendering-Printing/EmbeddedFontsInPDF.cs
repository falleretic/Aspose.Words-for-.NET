using Aspose.Words.Saving;

namespace Aspose.Words.Examples.CSharp.Rendering_and_Printing
{
    class EmbeddedFontsInPdf : TestDataHelper
    {
        public static void Run()
        {
            EmbeddedAllFonts();
            EmbeddedSubsetFonts();
        }

        private static void EmbeddedAllFonts()
        {
            //ExStart:EmbeddAllFonts
            Document doc = new Document(MailMergeDir + "Rendering.doc");

            // Aspose.Words embeds full fonts by default when EmbedFullFonts is set to true. The property below can be changed
            // Each time a document is rendered
            PdfSaveOptions options = new PdfSaveOptions();
            options.EmbedFullFonts = true;

            // The output PDF will be embedded with all fonts found in the document
            doc.Save(ArtifactsDir + "EmbeddedFontsInPdf.pdf", options);
            //ExEnd:EmbeddAllFonts
        }

        private static void EmbeddedSubsetFonts()
        {
            //ExStart:EmbeddSubsetFonts
            Document doc = new Document(MailMergeDir + "Rendering.doc");
            
            // To subset fonts in the output PDF document, simply create new PdfSaveOptions and set EmbedFullFonts to false
            PdfSaveOptions options = new PdfSaveOptions();
            options.EmbedFullFonts = false;
            
            // The output PDF will contain subsets of the fonts in the document. Only the glyphs used
            // in the document are included in the PDF fonts
            doc.Save(ArtifactsDir + "EmbeddSubsetFonts.pdf", options);
            //ExEnd:EmbeddSubsetFonts
        }
    }
}