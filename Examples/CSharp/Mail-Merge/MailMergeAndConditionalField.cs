namespace Aspose.Words.Examples.CSharp.Mail_Merge
{
    class MailMergeAndConditionalField : TestDataHelper
    {
        public static void Run()
        {
            //ExStart:MailMergeAndConditionalField
            Document doc = new Document(MailMergeDir + "UnconditionalMergeFieldsAndRegions.docx");

            // Merge fields and merge regions are merged regardless of the parent IF field's condition
            doc.MailMerge.UnconditionalMergeFieldsAndRegions = true;

            // Fill the fields in the document with user data
            doc.MailMerge.Execute(
                new[] { "FullName" },
                new object[] { "James Bond" });

            doc.Save(ArtifactsDir + "UnconditionalMergeFieldsAndRegions.docx");
            //ExEnd:MailMergeAndConditionalField
        }
    }
}