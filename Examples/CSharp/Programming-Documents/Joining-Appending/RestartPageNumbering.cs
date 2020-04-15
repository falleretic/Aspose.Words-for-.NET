using NUnit.Framework;

namespace Aspose.Words.Examples.CSharp.Programming_Documents.Joining_and_Appending
{
    class RestartPageNumbering : TestDataHelper
    {
        [Test]
        public static void Run()
        {
            //ExStart:RestartPageNumbering
            Document dstDoc = new Document(JoiningAppendingDir + "TestFile.Destination.doc");
            Document srcDoc = new Document(JoiningAppendingDir + "TestFile.Source.doc");

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