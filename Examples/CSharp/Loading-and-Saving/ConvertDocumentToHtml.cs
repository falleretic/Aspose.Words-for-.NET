using Aspose.Words.Saving;
using NUnit.Framework;

namespace Aspose.Words.Examples.CSharp.Loading_Saving
{
    class ConvertDocumentToHtml : TestDataHelper
    {
        [Test]
        public static void ExportRoundtripInformation()
        {
            //ExStart:ConvertDocumentToHtmlWithRoundtrip
            // Load the document from disk.
            Document doc = new Document(LoadingSavingDir + "Test File (doc).docx");

            HtmlSaveOptions options = new HtmlSaveOptions();
            // HtmlSaveOptions.ExportRoundtripInformation property specifies
            // Whether to write the roundtrip information when saving to HTML, MHTML or EPUB.
            // Default value is true for HTML and false for MHTML and EPUB.
            options.ExportRoundtripInformation = true;
            
            doc.Save(ArtifactsDir + "ExportRoundtripInformation_out.html", options);
            //ExEnd:ConvertDocumentToHtmlWithRoundtrip
        }

        [Test]
        public static void SplitDocumentByHeadingsHtml()
        {
            //ExStart:SplitDocumentByHeadingsHtml
            // Open a Word document
            Document doc = new Document(LoadingSavingDir + "Test File (doc).docx");
 
            HtmlSaveOptions options = new HtmlSaveOptions();
            // Split a document into smaller parts, in this instance split by heading
            options.DocumentSplitCriteria = DocumentSplitCriteria.HeadingParagraph;
 
            // Save the output file
            doc.Save(ArtifactsDir + "SplitDocumentByHeadings_out.html", options);
            //ExEnd:SplitDocumentByHeadingsHtml
        }

        [Test]
        public static void SplitDocumentBySectionsHtml()
        {
            // Open a Word document
            Document doc = new Document(LoadingSavingDir + "Test File (doc).docx");
 
            //ExStart:SplitDocumentBySectionsHtml
            HtmlSaveOptions options = new HtmlSaveOptions();
            options.DocumentSplitCriteria = DocumentSplitCriteria.HeadingParagraph;
            //ExEnd:SplitDocumentBySectionsHtml
            
            // Save the output file
            doc.Save(ArtifactsDir + "SplitDocumentBySections_out.html", options);
        }
    }
}
