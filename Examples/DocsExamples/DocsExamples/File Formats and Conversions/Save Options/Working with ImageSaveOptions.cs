﻿using Aspose.Words;
using Aspose.Words.Saving;
using NUnit.Framework;

namespace DocsExamples.File_Formats_and_Conversions.Save_Options
{
    internal class WorkingWithImageSaveOptions : DocsExamplesBase
    {
        [Test]
        public void ExposeThresholdControlForTiffBinarization()
        {
            //ExStart:ExposeThresholdControlForTiffBinarization
            Document doc = new Document(MyDir + "Rendering.docx");

            ImageSaveOptions saveOptions = new ImageSaveOptions(SaveFormat.Tiff)
            {
                TiffCompression = TiffCompression.Ccitt3,
                ImageColorMode = ImageColorMode.Grayscale,
                TiffBinarizationMethod = ImageBinarizationMethod.FloydSteinbergDithering,
                ThresholdForFloydSteinbergDithering = 254
            };

            doc.Save(ArtifactsDir + "WorkingWithImageSaveOptions.ExposeThresholdControlForTiffBinarization.tiff", saveOptions);
            //ExEnd:ExposeThresholdControlForTiffBinarization
        }

        [Test]
        public void GetTiffPageRange()
        {
            //ExStart:GetTiffPageRange
            Document doc = new Document(MyDir + "Rendering.docx");
            //ExStart:SaveAsTIFF
            doc.Save(ArtifactsDir + "WorkingWithImageSaveOptions.MultipageTiff.tiff");
            //ExEnd:SaveAsTIFF
            
            //ExStart:SaveAsTIFFUsingImageSaveOptions
            ImageSaveOptions saveOptions = new ImageSaveOptions(SaveFormat.Tiff)
            {
                PageSet = new PageSet(new PageRange(0, 1)), TiffCompression = TiffCompression.Ccitt4, Resolution = 160
            };

            doc.Save(ArtifactsDir + "WorkingWithImageSaveOptions.GetTiffPageRange.tiff", saveOptions);
            //ExEnd:SaveAsTIFFUsingImageSaveOptions
            //ExEnd:GetTiffPageRange
        }

        [Test]
        public void Format1BppIndexed()
        {
            //ExStart:Format1BppIndexed
            Document doc = new Document(MyDir + "Rendering.docx");

            ImageSaveOptions saveOptions = new ImageSaveOptions(SaveFormat.Png)
            {
                PageSet = new PageSet(1),
                ImageColorMode = ImageColorMode.BlackAndWhite,
                PixelFormat = ImagePixelFormat.Format1bppIndexed
            };

            doc.Save(ArtifactsDir + "WorkingWithImageSaveOptions.Format1BppIndexed.Png", saveOptions);
            //ExEnd:Format1BppIndexed
        }

        [Test]
        public void GetJpegPageRange()
        {
            // ExStart:GetJpegPageRange
            Document doc = new Document(MyDir + "Rendering.docx");
            doc.Save(ArtifactsDir + "WorkingWithImageSaveOptions.JpegDefaultOptions.jpg");
            
            // Render the third page only and set the JPEG quality to 80%.
            // In this case we need to pass the desired SaveFormat to the ImageSaveOptions constructor
            // to signal what type of image to save as.
            ImageSaveOptions saveOptions = new ImageSaveOptions(SaveFormat.Jpeg)
            {
                PageSet = new PageSet(2), JpegQuality = 80
            };

            doc.Save(ArtifactsDir + "WorkingWithImageSaveOptions.GetJpegPageRange.jpg", saveOptions);
            // ExEnd:GetJpegPageRange
        }
    }
}