﻿// Copyright (c) 2001-2020 Aspose Pty Ltd. All Rights Reserved.
//
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////

using System.Drawing;
using Aspose.Words;
using Aspose.Words.Saving;
using NUnit.Framework;
#if NETFRAMEWORK
using System.Drawing.Drawing2D;
using System.Drawing.Text;
#else
using SkiaSharp;
#endif

namespace ApiExamples
{
    [TestFixture]
    internal class ExImageSaveOptions : ApiExampleBase
    {
        [Test]
        public void UseGdiEmfRenderer()
        {
            //ExStart
            //ExFor:ImageSaveOptions.UseGdiEmfRenderer
            //ExSummary:Shows how to save metafiles directly without using GDI+ to EMF.
            Document doc = new Document(MyDir + "Rendering.docx");

            ImageSaveOptions saveOptions = new ImageSaveOptions(SaveFormat.Emf)
            {
                UseGdiEmfRenderer = false
            };

            doc.Save(ArtifactsDir + "ImageSaveOptions.UseGdiEmfRenderer.docx", saveOptions);
            //ExEnd
        }

        [Test]
        public void SaveIntoGif()
        {
            //ExStart
            //ExFor:ImageSaveOptions.PageIndex
            //ExSummary:Shows how to save specific document page as image file.
            Document doc = new Document(MyDir + "Rendering.docx");

            ImageSaveOptions saveOptions = new ImageSaveOptions(SaveFormat.Gif)
            {
                PageIndex = 1 // Define which page will save
            };

            doc.Save(ArtifactsDir + "ImageSaveOptions.SaveIntoGif.gif", saveOptions);
            //ExEnd
        }

#if NETFRAMEWORK
        [Test]
        public void GraphicsQuality()
        {
            //ExStart
            //ExFor:GraphicsQualityOptions
            //ExFor:GraphicsQualityOptions.CompositingMode
            //ExFor:GraphicsQualityOptions.CompositingQuality
            //ExFor:GraphicsQualityOptions.InterpolationMode
            //ExFor:GraphicsQualityOptions.StringFormat
            //ExFor:GraphicsQualityOptions.SmoothingMode
            //ExFor:GraphicsQualityOptions.TextRenderingHint
            //ExFor:ImageSaveOptions.GraphicsQualityOptions
            //ExSummary:Shows how to set render quality options when converting documents to image formats. 
            Document doc = new Document(MyDir + "Rendering.docx");

            GraphicsQualityOptions qualityOptions = new GraphicsQualityOptions
            {
                SmoothingMode = SmoothingMode.AntiAlias,
                TextRenderingHint = TextRenderingHint.ClearTypeGridFit,
                CompositingMode = CompositingMode.SourceOver,
                CompositingQuality = CompositingQuality.HighQuality,
                InterpolationMode = InterpolationMode.High,
                StringFormat = StringFormat.GenericTypographic
            };

            ImageSaveOptions saveOptions = new ImageSaveOptions(SaveFormat.Jpeg);
            saveOptions.GraphicsQualityOptions = qualityOptions;

            doc.Save(ArtifactsDir + "ImageSaveOptions.GraphicsQuality.jpeg", saveOptions);
            //ExEnd
        }

        [Test]
        public void WindowsMetaFile()
        {
            //ExStart
            //ExFor:ImageSaveOptions.MetafileRenderingOptions
            //ExSummary:Shows how to set the rendering mode for Windows Metafiles. 
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Use a DocumentBuilder to insert a .wmf image into the document
            builder.InsertImage(Image.FromFile(ImageDir + "Windows MetaFile.wmf"));

            // For documents that contain .wmf images, when converting the documents themselves to images,
            // we can use a ImageSaveOptions object to designate a rendering method for the .wmf images
            ImageSaveOptions options = new ImageSaveOptions(SaveFormat.Png);
            options.MetafileRenderingOptions.RenderingMode = MetafileRenderingMode.Bitmap;

            doc.Save(ArtifactsDir + "ImageSaveOptions.WindowsMetaFile.png", options);
            //ExEnd
        }
#endif

        [Test]
        [Category("SkipMono")]
        public void BlackAndWhite()
        {
            //ExStart
            //ExFor:ImageColorMode
            //ExFor:ImagePixelFormat
            //ExFor:ImageSaveOptions.Clone
            //ExFor:ImageSaveOptions.ImageColorMode
            //ExFor:ImageSaveOptions.PixelFormat
            //ExSummary:Show how to convert document images to black and white with 1 bit per pixel
            Document doc = new Document(MyDir + "Rendering.docx");

            ImageSaveOptions imageSaveOptions = new ImageSaveOptions(SaveFormat.Png);
            imageSaveOptions.ImageColorMode = ImageColorMode.BlackAndWhite;
            imageSaveOptions.PixelFormat = ImagePixelFormat.Format1bppIndexed;

            // ImageSaveOptions instances can be cloned
            Assert.AreNotEqual(imageSaveOptions, imageSaveOptions.Clone());  

            doc.Save(ArtifactsDir + "ImageSaveOptions.BlackAndWhite.png", imageSaveOptions);
            //ExEnd
        }

        [Test]
        public void FloydSteinbergDithering()
        {
            //ExStart
            //ExFor:ImageBinarizationMethod
            //ExFor:ImageSaveOptions.ThresholdForFloydSteinbergDithering
            //ExFor:ImageSaveOptions.TiffBinarizationMethod
            //ExSummary: Shows how to control the threshold for TIFF binarization in the Floyd-Steinberg method
            Document doc = new Document (MyDir + "Rendering.docx");

            ImageSaveOptions options = new ImageSaveOptions(SaveFormat.Tiff)
            {
                TiffCompression = TiffCompression.Ccitt3,
                ImageColorMode = ImageColorMode.Grayscale,
                TiffBinarizationMethod = ImageBinarizationMethod.FloydSteinbergDithering,
                // The default value of this property is 128. The higher value, the darker image
                ThresholdForFloydSteinbergDithering = 254
            };

            doc.Save(ArtifactsDir + "ImageSaveOptions.FloydSteinbergDithering.tiff", options);
            //ExEnd
        }

        [Test]
        public void EditImage()
        {
            //ExStart
            //ExFor:ImageSaveOptions.HorizontalResolution
            //ExFor:ImageSaveOptions.ImageBrightness
            //ExFor:ImageSaveOptions.ImageContrast
            //ExFor:ImageSaveOptions.SaveFormat
            //ExFor:ImageSaveOptions.Scale
            //ExFor:ImageSaveOptions.VerticalResolution
            //ExSummary:Shows how to edit image.
            Document doc = new Document(MyDir + "Rendering.docx");

            // When saving the document as an image, we can use an ImageSaveOptions object to edit various aspects of it
            ImageSaveOptions options = new ImageSaveOptions(SaveFormat.Png)
            {
                ImageBrightness = 0.3f,     // 0 - 1 scale, default at 0.5
                ImageContrast = 0.7f,       // 0 - 1 scale, default at 0.5
                HorizontalResolution = 72f, // Default at 96.0 meaning 96dpi, image dimensions will be affected if we change resolution
                VerticalResolution = 72f,   // Default at 96.0 meaning 96dpi
                Scale = 96f / 72f           // Default at 1.0 for normal scale, can be used to negate resolution impact in image size
            };

            doc.Save(ArtifactsDir + "ImageSaveOptions.EditImage.png", options);
            //ExEnd
        }
    }
}