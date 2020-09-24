using NUnit.Framework;

namespace Aspose.Words.Examples.CSharp
{
    class RestartPageNumbering : TestDataHelper
    {
        [Test]
        public static void Run()
        {
            //ExStart:RestartPageNumbering
            Document srcDoc = new Document(JoiningAppendingDir + "Document source.docx");
            Document dstDoc = new Document(JoiningAppendingDir + "Northwind traders.docx");

            // Set the appended document to appear on the next page
            srcDoc.FirstSection.PageSetup.SectionStart = SectionStart.NewPage;
            // Restart the page numbering for the document to be appended
            srcDoc.FirstSection.PageSetup.RestartPageNumbering = true;

            dstDoc.AppendDocument(srcDoc, ImportFormatMode.KeepSourceFormatting);
            
            dstDoc.Save(ArtifactsDir + "RestartPageNumbering.docx");
            //ExEnd:RestartPageNumbering
        }
    }
}