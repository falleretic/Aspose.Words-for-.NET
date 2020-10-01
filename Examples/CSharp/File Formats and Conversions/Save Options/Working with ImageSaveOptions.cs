using System.IO;
using Aspose.Words.Saving;
using NUnit.Framework;

namespace Aspose.Words.Examples.CSharp.File_Formats_and_Conversions.Save_Options
{
    class ImageColorFilters : TestDataHelper
    {
        [Test]
        public static void TiffCompressionLzw()
        {
            Document doc = new Document(MyDir + "Colors.docx");

            ImageSaveOptions imgOpttiff = new ImageSaveOptions(SaveFormat.Tiff);
            imgOpttiff.Resolution = 100;
            imgOpttiff.TiffCompression = TiffCompression.Lzw;
            imgOpttiff.ImageBrightness = 0.8f;
            imgOpttiff.ImageContrast = 0.8f;

            doc.Save(ArtifactsDir + "ImageSaveOptions.TiffCompressionLzw.tiff", imgOpttiff);
        }

        [Test]
        public static void GrayscaleTiffCompressionLzw()
        {
            Document doc = new Document(MyDir + "Colors.docx");

            ImageSaveOptions imgOpttiff = new ImageSaveOptions(SaveFormat.Tiff);
            imgOpttiff.Resolution = 100;
            imgOpttiff.TiffCompression = TiffCompression.Lzw;
            imgOpttiff.ImageColorMode = ImageColorMode.Grayscale;
            imgOpttiff.ImageBrightness = 0.8f;
            imgOpttiff.ImageContrast = 0.8f;

            doc.Save(ArtifactsDir + "ImageSaveOptions.GrayscaleTiffCompressionLzw.tiff", imgOpttiff);
        }

        [Test]
        public static void BlackWhiteTiffCompressionLzw()
        {
            Document doc = new Document(MyDir + "Colors.docx");

            ImageSaveOptions imgOpttiff = new ImageSaveOptions(SaveFormat.Tiff);
            imgOpttiff.Resolution = 100;
            imgOpttiff.TiffCompression = TiffCompression.Lzw;
            imgOpttiff.ImageColorMode = ImageColorMode.BlackAndWhite;
            // Set brightness and contrast according to high sensitivity.
            imgOpttiff.ImageBrightness = 0.4f;
            imgOpttiff.ImageContrast = 0.3f;
            
            doc.Save(ArtifactsDir + "ImageSaveOptions.BlackWhiteTiffCompressionLzw.tiff", imgOpttiff);
        }

        [Test]
        public static void BlackWhiteTiffCompressionCcitt4()
        {
            Document doc = new Document(MyDir + "Colors.docx");

            ImageSaveOptions imgOpttiff = new ImageSaveOptions(SaveFormat.Tiff);
            imgOpttiff.Resolution = 100;
            imgOpttiff.TiffCompression = TiffCompression.Ccitt4;
            imgOpttiff.ImageColorMode = ImageColorMode.Grayscale;
            // Set brightness and contrast according to high sensitivity.
            imgOpttiff.ImageBrightness = 0.4f;
            imgOpttiff.ImageContrast = 0.3f;
            
            doc.Save(ArtifactsDir + "ImageSaveOptions.BlackWhiteTiffCompressionCcitt4.tiff", imgOpttiff);
        }

        [Test]
        public static void BlackWhiteTiffCompressionRle()
        {
            Document doc = new Document(MyDir + "Colors.docx");

            ImageSaveOptions imgOpttiff = new ImageSaveOptions(SaveFormat.Tiff);
            imgOpttiff.Resolution = 100;
            imgOpttiff.TiffCompression = TiffCompression.Rle;
            imgOpttiff.ImageColorMode = ImageColorMode.Grayscale;
            // Set brightness and contrast according to high sensitivity.
            imgOpttiff.ImageBrightness = 0.4f;
            imgOpttiff.ImageContrast = 0.3f;
            
            doc.Save(ArtifactsDir + "ImageSaveOptions.BlackWhiteTiffCompressionRle.tiff", imgOpttiff);
        }

        [Test]
        public static void Format1bppIndexed()
        {
            //ExStart:Format1bppIndexed
            Document doc = new Document(MyDir + "Colors.docx");

            ImageSaveOptions opt = new ImageSaveOptions(SaveFormat.Png);
            opt.PageIndex = 1;
            opt.ImageColorMode = ImageColorMode.BlackAndWhite;
            opt.PixelFormat = ImagePixelFormat.Format1bppIndexed;

            doc.Save(ArtifactsDir + "ImageSaveOptions.Format1bppIndexed.Png", opt);
            //ExEnd:Format1bppIndexed
        }

        [Test]
        public static void ExposeThresholdControlForTiffBinarization()
        {
            //ExStart:ExposeThresholdControlForTiffBinarization
            Document doc = new Document(MyDir + "Colors.docx");

            ImageSaveOptions options = new ImageSaveOptions(SaveFormat.Tiff);
            options.TiffCompression = TiffCompression.Ccitt3;
            options.ImageColorMode = ImageColorMode.Grayscale;
            options.TiffBinarizationMethod = ImageBinarizationMethod.FloydSteinbergDithering;
            options.ThresholdForFloydSteinbergDithering = 254;

            doc.Save(ArtifactsDir + "ImageSaveOptions.ExposeThresholdControlForTiffBinarization.tiff", options);
            //ExEnd:ExposeThresholdControlForTiffBinarization
        }

        [Test]
        public static void SaveDocumentToJpeg()
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

        [Test]
        public static void SaveAsMultipageTiff()
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
    }
}