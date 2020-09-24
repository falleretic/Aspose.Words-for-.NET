using Aspose.Words.MailMerging;
using NUnit.Framework;

namespace Aspose.Words.Examples.CSharp
{
    class MailMergeCleanUp : TestDataHelper
    {
        [Test]
        public static void CleanupParagraphsWithPunctuationMarks()
        {
            //ExStart:CleanupParagraphsWithPunctuationMarks
            Document doc = new Document(MailMergeDir + "Mail merge destinations - Cleanup punctuation marks.docx");

            doc.MailMerge.CleanupOptions = MailMergeCleanupOptions.RemoveEmptyParagraphs;
            doc.MailMerge.CleanupParagraphsWithPunctuationMarks = false;

            doc.MailMerge.Execute(new string[] { "field1", "field2" }, new object[] { "", "" });

            doc.Save(ArtifactsDir + "MailMerge.CleanupPunctuationMarks.docx");
            //ExEnd:CleanupParagraphsWithPunctuationMarks
        }
    }
}