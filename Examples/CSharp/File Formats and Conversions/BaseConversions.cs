using System;
using System.Drawing;
using System.Drawing.Imaging;
using System.IO;
using Aspose.Email;
using Aspose.Email.Clients.Smtp;
using Aspose.Words.Drawing;
using NUnit.Framework;

namespace Aspose.Words.Examples.CSharp
{
    class BaseConversions : TestDataHelper
    {
        [Test]
        public static void DocToDocx()
        {
            //ExStart:LoadAndSave
            //ExStart:OpenDocument
            Document doc = new Document(QuickStartDir + "Document.doc");
            //ExEnd:OpenDocument
            doc.Save(ArtifactsDir + "BaseConversions.DocToDocx.docx");
            //ExEnd:LoadAndSave
        }

        [Test]
        public static void DocxToRtf()
        {
            //ExStart:LoadAndSaveToStream 
            //ExStart:OpeningFromStream
            // Open the stream
            // Read only access is enough for Aspose.Words to load a document
            Stream stream = File.OpenRead(QuickStartDir + "Document.docx");

            Document doc = new Document(stream);
            // You can close the stream now, it is no longer needed because the document is in memory
            stream.Close();
            //ExEnd:OpeningFromStream 

            // ... do something with the document

            // Convert the document to a different format and save to stream
            MemoryStream dstStream = new MemoryStream();
            doc.Save(dstStream, SaveFormat.Rtf);

            // Rewind the stream position back to zero so it is ready for the next reader
            dstStream.Position = 0;
            //ExEnd:LoadAndSaveToStream 
            // Save the document from stream, to disk
            // Normally you would do something with the stream directly, for example writing the data to a database
            File.WriteAllBytes(ArtifactsDir + "BaseConversions.DocxToRtf.rtf", dstStream.ToArray());
        }

        [Test]
        public static void DocxToPdf()
        {
            //ExStart:Doc2Pdf
            Document doc = new Document(LoadingSavingDir + "Document.docx");
            doc.Save(ArtifactsDir + "BaseConversions.DocxToPdf.pdf");
            //ExEnd:Doc2Pdf
        }

        [Test]
        public static void DocxToByte()
        {
            //ExStart:ConvertDocumentToByte
            Document doc = new Document(LoadingSavingDir + "Document.docx");

            // Create a new memory stream
            MemoryStream outStream = new MemoryStream();
            // Save the document to stream
            doc.Save(outStream, SaveFormat.Docx);

            // Convert the document to byte form
            byte[] docBytes = outStream.ToArray();

            // The bytes are now ready to be stored/transmitted
            // Now reverse the steps to load the bytes back into a document object
            MemoryStream inStream = new MemoryStream(docBytes);

            // Load the stream into a new document object
            Document loadDoc = new Document(inStream);
            //ExEnd:ConvertDocumentToByte
        }

        [Test]
        public static void DocxToEpub()
        {
            //ExStart:ConvertDocumentToEPUBUsingDefaultSaveOption
            Document doc = new Document(LoadingSavingDir + "Document.docx");
            doc.Save(ArtifactsDir + "BaseConversions.DocxToEpub.epub");
            //ExEnd:ConvertDocumentToEPUBUsingDefaultSaveOption
        }

        [Test, Ignore("Only for example")]
        public static void DocxToMhtmlAndSendingEmail()
        {
            //ExStart:ConvertDocumentToMhtmlAndEmail
            Document doc = new Document(LoadingSavingDir + "Document.docx");

            // Save into a memory stream in MHTML format
            Stream stream = new MemoryStream();
            doc.Save(stream, SaveFormat.Mhtml);

            // Rewind the stream to the beginning so Aspose.Email can read it
            stream.Position = 0;

            // Create an Aspose.Network MIME email message from the stream
            MailMessage message = MailMessage.Load(stream, new MhtmlLoadOptions());
            message.From = "your_from@email.com";
            message.To = "your_to@email.com";
            message.Subject = "Aspose.Words + Aspose.Email MHTML Test Message";

            // Send the message using Aspose.Email
            SmtpClient client = new SmtpClient();
            client.Host = "your_smtp.com";
            client.Send(message);
            //ExEnd:ConvertDocumentToMhtmlAndEmail
        }

        [Test]
        public static void DocxToTxt()
        {
            //ExStart:SaveAsTxt
            Document doc = new Document(LoadingSavingDir + "Document.docx");
            doc.Save(ArtifactsDir + "BaseConversions.DocxToTxt.txt");
            //ExEnd:SaveAsTxt
        }

        [Test]
        public static void TxtToDocx()
        {
            //ExStart:LoadTxt
            // The encoding of the text file is automatically detected
            Document doc = new Document(LoadingSavingDir + "Txt document.txt");
            doc.Save(ArtifactsDir + "BaseConversions.TxtToDocx.docx");
            //ExEnd:LoadTxt
        }

        [Test]
        public static void ImagesToPdf()
        {
            //ExStart:ImageToPdf
            ConvertImageToPdf(LoadingSavingDir + "Test.jpg", ArtifactsDir + "BaseConversions.JpgToPdf.pdf");
            ConvertImageToPdf(LoadingSavingDir + "Test.png", ArtifactsDir + "BaseConversions.PngToPdf.pdf");
            ConvertImageToPdf(LoadingSavingDir + "Test.wmf", ArtifactsDir + "BaseConversions.WmfToPdf.pdf");
            ConvertImageToPdf(LoadingSavingDir + "Test.tiff", ArtifactsDir + "BaseConversions.TiffToPdf.pdf");
            ConvertImageToPdf(LoadingSavingDir + "Test.gif", ArtifactsDir + "BaseConversions.GifToPdf.pdf");
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