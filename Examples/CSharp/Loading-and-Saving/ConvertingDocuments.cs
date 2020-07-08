using System.IO;
using Aspose.Email;
using Aspose.Email.Clients.Smtp;
using NUnit.Framework;

namespace Aspose.Words.Examples.CSharp.Loading_Saving
{
    class ConvertingDocuments : TestDataHelper
    {
        [Test]
        public static void DocumentToByte()
        {
            //ExStart:ConvertDocumentToByte
            Document doc = new Document(LoadingSavingDir + "Test File (doc).doc");

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
        public static void DocumentToEpub()
        {
            //ExStart:ConvertDocumentToEPUBUsingDefaultSaveOption
            Document doc = new Document(LoadingSavingDir + "Document.EpubConversion.doc");
            doc.Save(ArtifactsDir + "ConvertDocumentToEPUBUsingDefaultSaveOption.epub");
            //ExEnd:ConvertDocumentToEPUBUsingDefaultSaveOption
        }

        [Test, Ignore("Only for example")]
        public static void DocumentToMhtmlAndSendingEmail()
        {
            //ExStart:ConvertDocumentToMhtmlAndEmail
            Document doc = new Document(LoadingSavingDir + "Test File (docx).docx");

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
        public static void SaveDoc2Pdf()
        {
            //ExStart:Doc2Pdf
            Document doc = new Document(LoadingSavingDir + "Rendering.doc");
            doc.Save(ArtifactsDir + "SaveDoc2Pdf.pdf");
            //ExEnd:Doc2Pdf
        }

        [Test]
        public static void ConvertToDocx()
        {
            //ExStart:LoadAndSave
            //ExStart:OpenDocument
            Document doc = new Document(QuickStartDir + "Document.doc");
            //ExEnd:OpenDocument
            doc.Save(ArtifactsDir + "LoadAndSaveToDisk.docx");
            //ExEnd:LoadAndSave
        }

        [Test]
        public static void ConvertToRtf()
        {
            //ExStart:LoadAndSaveToStream 
            //ExStart:OpeningFromStream
            // Open the stream
            // Read only access is enough for Aspose.Words to load a document
            Stream stream = File.OpenRead(QuickStartDir + "Document.doc");

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
            File.WriteAllBytes(ArtifactsDir + "LoadAndSaveToStream.rtf", dstStream.ToArray());
        }

        [Test]
        public static void ConvertTxtToDocx()
        {
            //ExStart:LoadTxt
            // The encoding of the text file is automatically detected
            Document doc = new Document(LoadingSavingDir + "LoadTxt.txt");
            doc.Save(ArtifactsDir + "LoadTxt.docx");
            //ExEnd:LoadTxt
        }
    }
}