using NUnit.Framework;

namespace Aspose.Words.Examples.CSharp
{
    class ExecuteArray : TestDataHelper
    {
        [Test]
        public static void Run()
        {
            //ExStart:ExecuteArray
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            builder.InsertField(" MERGEFIELD FullName ");
            builder.InsertParagraph();
            builder.InsertField(" MERGEFIELD Company ");
            builder.InsertParagraph();
            builder.InsertField(" MERGEFIELD Address ");
            builder.InsertParagraph();
            builder.InsertField(" MERGEFIELD City ");

            // Trim trailing and leading whitespaces mail merge values
            doc.MailMerge.TrimWhitespaces = false;

            // Fill the fields in the document with user data
            doc.MailMerge.Execute(new string[] { "FullName", "Company", "Address", "City" },
                new object[] { "James Bond", "MI5 Headquarters", "Milbank", "London" });

            // Send the document in Word format to the client browser with an option
            // to save to disk or open inside the current browser
            doc.Save(ArtifactsDir + "MailMerge.ExecuteArray.docx");
            //ExEnd:ExecuteArray
        }
    }
}