using System.IO;
using Aspose.Words.Saving;
using NUnit.Framework;

namespace Aspose.Words.Examples.CSharp
{
    class ImageColorFilters : TestDataHelper
    {
        [Test]
        public static void Run()
        {
            Document doc = new Document($"{RenderingPrintingDir}Colors.docx");

            SaveColorTiffWithLzw(doc, 0.8f, 0.8f);
            SaveGrayscaleTiffWithLzw(doc, 0.8f, 0.8f);
            SaveBlackWhiteTiffWithLzw(doc, true);
            SaveBlackWhiteTiffWithCitt4(doc, true);
            SaveBlackWhiteTiffWithRle(doc, true);
            SaveImageToOnebitPerPixel(doc);
            ExposeThresholdControlForTiffBinarization();
        }

        private static void SaveColorTiffWithLzw(Document doc, float brightness, float contrast)
        {
            // Select the TIFF format with 100 dpi
            ImageSaveOptions imgOpttiff = new ImageSaveOptions(SaveFormat.Tiff);
            imgOpttiff.Resolution = 100;

            // Select fullcolor LZW compression
            imgOpttiff.TiffCompression = TiffCompression.Lzw;

            // Set brightness and contrast
            imgOpttiff.ImageBrightness = brightness;
            imgOpttiff.ImageContrast = contrast;

            // Save multipage color TIFF
            doc.Save($"{ArtifactsDir}Result Colors.tiff", imgOpttiff);
        }

        private static void SaveGrayscaleTiffWithLzw(Document doc, float brightness, float contrast)
        {
            // Select the TIFF format with 100 dpi
            ImageSaveOptions imgOpttiff = new ImageSaveOptions(SaveFormat.Tiff);
            imgOpttiff.Resolution = 100;

            // Select LZW compression
            imgOpttiff.TiffCompression = TiffCompression.Lzw;

            // Apply grayscale filter
            imgOpttiff.ImageColorMode = ImageColorMode.Grayscale;

            // Set brightness and contrast
            imgOpttiff.ImageBrightness = brightness;
            imgOpttiff.ImageContrast = contrast;

            // Save multipage grayscale TIFF
            doc.Save($"{ArtifactsDir}Result Grayscale.tiff", imgOpttiff);
        }

        private static void SaveBlackWhiteTiffWithLzw(Document doc, bool highSensitivity)
        {
            // Select the TIFF format with 100 dpi
            ImageSaveOptions imgOpttiff = new ImageSaveOptions(SaveFormat.Tiff);
            imgOpttiff.Resolution = 100;

            // Apply black & white filter
            // Set very high sensitivity to gray color
            imgOpttiff.TiffCompression = TiffCompression.Lzw;
            imgOpttiff.ImageColorMode = ImageColorMode.BlackAndWhite;

            // Set brightness and contrast according to sensitivity
            if (highSensitivity)
            {
                imgOpttiff.ImageBrightness = 0.4f;
                imgOpttiff.ImageContrast = 0.3f;
            }
            else
            {
                imgOpttiff.ImageBrightness = 0.9f;
                imgOpttiff.ImageContrast = 0.9f;
            }

            // Save multipage TIFF
            doc.Save($"{ArtifactsDir}result black and white.tiff", imgOpttiff);
        }

        private static void SaveBlackWhiteTiffWithCitt4(Document doc, bool highSensitivity)
        {
            // Select the TIFF format with 100 dpi
            ImageSaveOptions imgOpttiff = new ImageSaveOptions(SaveFormat.Tiff);
            imgOpttiff.Resolution = 100;

            // Set CCITT4 compression
            imgOpttiff.TiffCompression = TiffCompression.Ccitt4;

            // Apply grayscale filter
            imgOpttiff.ImageColorMode = ImageColorMode.Grayscale;

            // Set brightness and contrast according to sensitivity
            if (highSensitivity)
            {
                imgOpttiff.ImageBrightness = 0.4f;
                imgOpttiff.ImageContrast = 0.3f;
            }
            else
            {
                imgOpttiff.ImageBrightness = 0.9f;
                imgOpttiff.ImageContrast = 0.9f;
            }

            // Save multipage TIFF
            doc.Save($"{ArtifactsDir}result Ccitt4.tiff", imgOpttiff);
        }

        private static void SaveBlackWhiteTiffWithRle(Document doc, bool highSensitivity)
        {
            // Select the TIFF format with 100 dpi
            ImageSaveOptions imgOpttiff = new ImageSaveOptions(SaveFormat.Tiff);
            imgOpttiff.Resolution = 100;

            // Set RLE compression
            imgOpttiff.TiffCompression = TiffCompression.Rle;

            // Aply grayscale filter
            imgOpttiff.ImageColorMode = ImageColorMode.Grayscale;

            // Set brightness and contrast according to sensitivity
            if (highSensitivity)
            {
                imgOpttiff.ImageBrightness = 0.4f;
                imgOpttiff.ImageContrast = 0.3f;
            }
            else
            {
                imgOpttiff.ImageBrightness = 0.9f;
                imgOpttiff.ImageContrast = 0.9f;
            }

            // Save multipage TIFF grayscale with low bright and contrast
            doc.Save($"{ArtifactsDir}result Rle.tiff", imgOpttiff);
        }

        private static void SaveImageToOnebitPerPixel(Document doc)
        {
            //ExStart:SaveImageToOnebitPerPixel
            ImageSaveOptions opt = new ImageSaveOptions(SaveFormat.Png);
            opt.PageIndex = 1;
            opt.ImageColorMode = ImageColorMode.BlackAndWhite;
            opt.PixelFormat = ImagePixelFormat.Format1bppIndexed;

            doc.Save(ArtifactsDir + "Format1bppIndexed.Png", opt);
            //ExEnd:SaveImageToOnebitPerPixel
        }

        private static void ExposeThresholdControlForTiffBinarization()
        {
            //ExStart:ExposeThresholdControlForTiffBinarization
            Document doc = new Document(RenderingPrintingDir + "Colors.docx");

            ImageSaveOptions options = new ImageSaveOptions(SaveFormat.Tiff);
            options.TiffCompression = TiffCompression.Ccitt3;
            options.ImageColorMode = ImageColorMode.Grayscale;
            options.TiffBinarizationMethod = ImageBinarizationMethod.FloydSteinbergDithering;
            options.ThresholdForFloydSteinbergDithering = 254;

            doc.Save(ArtifactsDir + "ThresholdForFloydSteinbergDithering.tiff", options);
            //ExEnd:ExposeThresholdControlForTiffBinarization
        }

        public static void SaveDocumentToJPEG()
        {
            // ExStart:SaveDocumentToJPEG
            Document doc = new Document(RenderingPrintingDir + "Rendering.docx");
            // Save as a JPEG image file with default options
            doc.Save(ArtifactsDir + "Rendering.JpegDefaultOptions.jpg");

            // Save document to stream as a JPEG with default options
            MemoryStream docStream = new MemoryStream();
            doc.Save(docStream, SaveFormat.Jpeg);
            // Rewind the stream position back to the beginning, ready for use
            docStream.Seek(0, SeekOrigin.Begin);

            // Save document to a JPEG image with specified options.
            // Render the third page only and set the JPEG quality to 80%
            // In this case we need to pass the desired SaveFormat to the ImageSaveOptions constructor 
            // to signal what type of image to save as.
            ImageSaveOptions imageOptions = new ImageSaveOptions(SaveFormat.Jpeg);
            imageOptions.PageIndex = 2;
            imageOptions.PageCount = 1;
            imageOptions.JpegQuality = 80;
            doc.Save(ArtifactsDir + "Rendering.JpegCustomOptions.jpg", imageOptions);
            // ExEnd:SaveDocumentToJPEG
        }

        [Test]
        public static void SaveAsMultipageTiff()
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