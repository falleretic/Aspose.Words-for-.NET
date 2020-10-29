using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;
using NUnit.Framework;

namespace SiteExamples.File_Formats_and_Conversions.Save_Options
{
    internal class ImageColorFilters : SiteExamplesBase
    {
        [Test, Description("Shows how to use image optimization when saving to TIFF.")]
        public void ExposeThresholdControlForTiffBinarization()
        {
            //ExStart:ExposeThresholdControlForTiffBinarization
            Document doc = new Document(MyDir + "Rendering.docx");

            ImageSaveOptions options = new ImageSaveOptions(SaveFormat.Tiff);
            options.TiffCompression = TiffCompression.Ccitt3;
            options.ImageColorMode = ImageColorMode.Grayscale;
            options.TiffBinarizationMethod = ImageBinarizationMethod.FloydSteinbergDithering;
            options.ThresholdForFloydSteinbergDithering = 254;

            doc.Save(ArtifactsDir + "ImageSaveOptions.ExposeThresholdControlForTiffBinarization.tiff", options);
            //ExEnd:ExposeThresholdControlForTiffBinarization
        }

        [Test, Description("Shows how to save several pages to TIFF.")]
        public void SaveAsMultipageTiff()
        {
            //ExStart:SaveAsMultipageTiff
            Document doc = new Document(MyDir + "Rendering.docx");
            //ExStart:SaveAsTIFF
            doc.Save(ArtifactsDir + "ImageSaveOptions.MultipageTIFF.tiff");
            //ExEnd:SaveAsTIFF
            
            //ExStart:SaveAsTIFFUsingImageSaveOptions
            ImageSaveOptions options = new ImageSaveOptions(SaveFormat.Tiff);
            options.PageIndex = 0;
            options.PageCount = 2;
            options.TiffCompression = TiffCompression.Ccitt4;
            options.Resolution = 160;
            
            doc.Save(ArtifactsDir + "ImageSaveOptions.SaveAsMultipageTiff.tiff", options);
            //ExEnd:SaveAsTIFFUsingImageSaveOptions
            //ExEnd:SaveAsMultipageTiff
        }

        [Test, Description("Shows how to set pixel format for the images.")]
        public void Format1bppIndexed()
        {
            //ExStart:Format1bppIndexed
            Document doc = new Document(MyDir + "Rendering.docx");

            ImageSaveOptions opt = new ImageSaveOptions(SaveFormat.Png);
            opt.PageIndex = 1;
            opt.ImageColorMode = ImageColorMode.BlackAndWhite;
            opt.PixelFormat = ImagePixelFormat.Format1bppIndexed;

            doc.Save(ArtifactsDir + "ImageSaveOptions.Format1bppIndexed.Png", opt);
            //ExEnd:Format1bppIndexed
        }

        [Test, Description("Shows how to set quality of the JPEG image.")]
        public void SaveDocumentToJpeg()
        {
            // ExStart:SaveDocumentToJpeg
            Document doc = new Document(MyDir + "Rendering.docx");
            doc.Save(ArtifactsDir + "ImageSaveOptions.JpegDefaultOptions.jpg");

            MemoryStream docStream = new MemoryStream();
            doc.Save(docStream, SaveFormat.Jpeg);
            // Rewind the stream position back to the beginning, ready for use.
            docStream.Seek(0, SeekOrigin.Begin);

            // Save document to a JPEG image with specified options.
            // Render the third page only and set the JPEG quality to 80%.
            // In this case we need to pass the desired SaveFormat to the ImageSaveOptions constructor
            // to signal what type of image to save as.
            ImageSaveOptions imageOptions = new ImageSaveOptions(SaveFormat.Jpeg);
            imageOptions.PageIndex = 2;
            imageOptions.PageCount = 1;
            imageOptions.JpegQuality = 80;
            doc.Save(ArtifactsDir + "ImageSaveOptions.SaveDocumentToJpeg.jpg", imageOptions);
            // ExEnd:SaveDocumentToJpeg
        }
    }
}