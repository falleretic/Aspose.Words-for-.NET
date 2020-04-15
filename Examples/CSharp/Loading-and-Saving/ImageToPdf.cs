﻿using System;
using System.Drawing;
using System.Drawing.Imaging;
using Aspose.Words.Drawing;
using NUnit.Framework;

namespace Aspose.Words.Examples.CSharp.Loading_Saving
{
    class ImageToPdf : TestDataHelper
    {
        [Test]
        public static void Run()
        {
            //ExStart:ImageToPdf
            ConvertImageToPdf(LoadingSavingDir + "Test.jpg", ArtifactsDir + "TestJpg.pdf");
            ConvertImageToPdf(LoadingSavingDir + "Test.png", ArtifactsDir + "TestPng.pdf");
            ConvertImageToPdf(LoadingSavingDir + "Test.wmf", ArtifactsDir + "TestWmf.pdf");
            ConvertImageToPdf(LoadingSavingDir + "Test.tiff", ArtifactsDir + "TestTif.pdf");
            ConvertImageToPdf(LoadingSavingDir + "Test.gif", ArtifactsDir + "TestGif.pdf");
            //ExEnd:ImageToPdf
        }

        /// <summary>
        /// Converts an image to PDF using Aspose.Words for .NET.
        /// </summary>
        /// <param name="inputFileName">File name of input image file.</param>
        /// <param name="outputFileName">Output PDF file name.</param>
        public static void ConvertImageToPdf(string inputFileName, string outputFileName)
        {
            Console.WriteLine("Converting " + inputFileName + " to PDF ....");

            //ExStart:ConvertImageToPdf
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Read the image from file, ensure it is disposed
            using (Image image = Image.FromFile(inputFileName))
            {
                // Find which dimension the frames in this image represent. For example 
                // The frames of a BMP or TIFF are "page dimension" whereas frames of a GIF image are "time dimension"
                FrameDimension dimension = new FrameDimension(image.FrameDimensionsList[0]);

                // Get the number of frames in the image
                int framesCount = image.GetFrameCount(dimension);

                // Loop through all frames
                for (int frameIdx = 0; frameIdx < framesCount; frameIdx++)
                {
                    // Insert a section break before each new page, in case of a multi-frame TIFF
                    if (frameIdx != 0)
                        builder.InsertBreak(BreakType.SectionBreakNewPage);

                    // Select active frame
                    image.SelectActiveFrame(dimension, frameIdx);

                    // We want the size of the page to be the same as the size of the image
                    // Convert pixels to points to size the page to the actual image size
                    PageSetup ps = builder.PageSetup;
                    ps.PageWidth = ConvertUtil.PixelToPoint(image.Width, image.HorizontalResolution);
                    ps.PageHeight = ConvertUtil.PixelToPoint(image.Height, image.VerticalResolution);

                    // Insert the image into the document and position it at the top left corner of the page
                    builder.InsertImage(
                        image,
                        RelativeHorizontalPosition.Page,
                        0,
                        RelativeVerticalPosition.Page,
                        0,
                        ps.PageWidth,
                        ps.PageHeight,
                        WrapType.None);
                }
            }

            doc.Save(outputFileName);
            //ExEnd:ConvertImageToPdf
        }
    }
}