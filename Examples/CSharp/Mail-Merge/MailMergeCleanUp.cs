using Aspose.Words.MailMerging;

namespace Aspose.Words.Examples.CSharp.Mail_Merge
{
    class MailMergeCleanUp : TestDataHelper
    {
        public static void Run()
        {
            CleanupParagraphsWithPunctuationMarks();
        }

        public static void CleanupParagraphsWithPunctuationMarks()
        {
            //ExStart:CleanupParagraphsWithPunctuationMarks
            Document doc = new Document(MailMergeDir + "MailMerge.CleanupPunctuationMarks.docx");

            doc.MailMerge.CleanupOptions = MailMergeCleanupOptions.RemoveEmptyParagraphs;
            doc.MailMerge.CleanupParagraphsWithPunctuationMarks = false;

            doc.MailMerge.Execute(new string[] { "field1", "field2" }, new object[] { "", "" });

            doc.Save(ArtifactsDir + "MailMerge.CleanupPunctuationMarks.docx");
            //ExEnd:CleanupParagraphsWithPunctuationMarks
        }
    }
}