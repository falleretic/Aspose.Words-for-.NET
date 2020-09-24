using NUnit.Framework;

namespace Aspose.Words.Examples.CSharp
{
    class JoinNewPage : TestDataHelper
    {
        [Test]
        public static void Run()
        {
            //ExStart:JoinNewPage
            Document srcDoc = new Document(JoiningAppendingDir + "Document source.docx");
            Document dstDoc = new Document(JoiningAppendingDir + "Northwind traders.docx");

            // Set the appended document to start on a new page
            srcDoc.FirstSection.PageSetup.SectionStart = SectionStart.NewPage;

            // Append the source document using the original styles found in the source document
            dstDoc.AppendDocument(srcDoc, ImportFormatMode.KeepSourceFormatting);
            
            dstDoc.Save(ArtifactsDir + "JoinNewPage.docx");
            //ExEnd:JoinNewPage
        }
    }
}