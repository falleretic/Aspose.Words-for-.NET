using System.IO;
using Aspose.Email;
using Aspose.Email.Clients.Smtp;
using NUnit.Framework;

namespace Aspose.Words.Examples.CSharp.Loading_and_Saving
{
    class ConvertDocumentToMhtmlAndEmail : TestDataHelper
    {
        [Test, Ignore("Only for example")]
        public static void Run()
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
    }
}