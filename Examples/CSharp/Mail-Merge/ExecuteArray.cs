namespace Aspose.Words.Examples.CSharp.Mail_Merge
{
    class ExecuteArray : TestDataHelper
    {
        public static void Run()
        {
            //ExStart:ExecuteArray
            Document doc = new Document(MailMergeDir + "MailMerge.ExecuteArray.doc");

            // Trim trailing and leading whitespaces mail merge values
            doc.MailMerge.TrimWhitespaces = false;

            // Fill the fields in the document with user data
            doc.MailMerge.Execute(
                new string[] { "FullName", "Company", "Address", "Address2", "City" },
                new object[] { "James Bond", "MI5 Headquarters", "Milbank", "", "London" });

            // Send the document in Word format to the client browser with an option
            // to save to disk or open inside the current browser
            doc.Save(ArtifactsDir + "MailMerge.ExecuteArray.doc");
            //ExEnd:ExecuteArray
        }
    }
}