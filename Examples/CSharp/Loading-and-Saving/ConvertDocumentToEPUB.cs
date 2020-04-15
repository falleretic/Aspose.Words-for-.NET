using Aspose.Words.Saving;
using NUnit.Framework;

namespace Aspose.Words.Examples.CSharp.Loading_Saving
{
    class ConvertDocumentToEPUB : TestDataHelper
    {
        [Test]
        public static void Run()
        {
            //ExStart:ConvertDocumentToEPUB
            Document doc = new Document(LoadingSavingDir + "Document.EpubConversion.doc");

            // Create a new instance of HtmlSaveOptions. This object allows us to set options that control
            // how the output document is saved
            HtmlSaveOptions saveOptions = new HtmlSaveOptions();
            // Specify the desired encoding
            saveOptions.Encoding = System.Text.Encoding.UTF8;
            // Specify at what elements to split the internal HTML at. This creates a new HTML within the EPUB 
            // which allows you to limit the size of each HTML part. This is useful for readers which cannot read 
            // HTML files greater than a certain size e.g 300kb.
            saveOptions.DocumentSplitCriteria = DocumentSplitCriteria.HeadingParagraph;
            // Specify that we want to export document properties
            saveOptions.ExportDocumentProperties = true;
            // Specify that we want to save in EPUB format
            saveOptions.SaveFormat = SaveFormat.Epub;

            // Export the document as an EPUB file
            doc.Save(ArtifactsDir + "ConvertDocumentToEPUB.epub", saveOptions);
            //ExEnd:ConvertDocumentToEPUB
        }

        [Test]
        public static void ConvertDocumentToEpubUsingDefaultSaveOption()
        {
            //ExStart:ConvertDocumentToEPUBUsingDefaultSaveOption
            Document doc = new Document(LoadingSavingDir + "Document.EpubConversion.doc");
            doc.Save(ArtifactsDir + "ConvertDocumentToEPUBUsingDefaultSaveOption.epub");
            //ExEnd:ConvertDocumentToEPUBUsingDefaultSaveOption
        }
    }
}