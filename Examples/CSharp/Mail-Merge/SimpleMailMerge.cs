﻿namespace Aspose.Words.Examples.CSharp.Mail_Merge
{
    class SimpleMailMerge : TestDataHelper
    {
        public static void Run()
        {
            //ExStart:SimpleMailMerge
            Document doc = new Document(MailMergeDir + "MailMerge.ExecuteArray.doc");

            doc.MailMerge.UseNonMergeFields = true;

            // Fill the fields in the document with user data
            doc.MailMerge.Execute(
                new string[] { "FullName", "Company", "Address", "Address2", "City" },
                new object[] { "James Bond", "MI5 Headquarters", "Milbank", "", "London" });

            // Send the document in Word format to the client browser with an option to save to disk or open inside the current browser
            doc.Save(ArtifactsDir + "SimpleMailMerge.docx");
            //ExEnd:SimpleMailMerge
        }
    }
}