using Aspose.Words.Saving;
using NUnit.Framework;

namespace Aspose.Words.Examples.CSharp
{
    class SaveAsMultipageTiff : TestDataHelper
    {
        [Test]
        public static void Run()
        {
            //ExStart:SaveAsMultipageTiff
            Document doc = new Document(RenderingPrintingDir + "Rendering.docx");
            //ExStart:SaveAsTIFF
            // Save the document as multipage TIFF
            doc.Save(ArtifactsDir + "MultipageTIFF.tiff");
            //ExEnd:SaveAsTIFF
            
            //ExStart:SaveAsTIFFUsingImageSaveOptions
            // Create an ImageSaveOptions object to pass to the Save method
            ImageSaveOptions options = new ImageSaveOptions(SaveFormat.Tiff);
            options.PageIndex = 0;
            options.PageCount = 2;
            options.TiffCompression = TiffCompression.Ccitt4;
            options.Resolution = 160;
            
            doc.Save(ArtifactsDir + "SaveAsMultipageTiff.tiff", options);
            //ExEnd:SaveAsTIFFUsingImageSaveOptions
            //ExEnd:SaveAsMultipageTiff
        }
    }
}