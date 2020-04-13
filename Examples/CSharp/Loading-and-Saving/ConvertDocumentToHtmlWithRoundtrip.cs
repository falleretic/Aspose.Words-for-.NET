using Aspose.Words.Saving;

namespace Aspose.Words.Examples.CSharp.Loading_Saving
{
    class ConvertDocumentToHtmlWithRoundtrip : TestDataHelper
    {
        public static void Run()
        {
            //ExStart:ConvertDocumentToHtmlWithRoundtrip
            Document doc = new Document(LoadingSavingDir + "Test File (doc).doc");

            HtmlSaveOptions options = new HtmlSaveOptions();
            // HtmlSaveOptions.ExportRoundtripInformation property specifies
            // Whether to write the roundtrip information when saving to HTML, MHTML or EPUB
            // Default value is true for HTML and false for MHTML and EPUB
            options.ExportRoundtripInformation = true;

            doc.Save(ArtifactsDir + "ConvertDocumentToHtmlWithRoundtrip.html", options);
            //ExEnd:ConvertDocumentToHtmlWithRoundtrip
        }
    }
}