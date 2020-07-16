using NUnit.Framework;

namespace Aspose.Words.Examples.CSharp
{
    class LinkHeadersFooters : TestDataHelper
    {
        [Test]
        public static void Run()
        {
            //ExStart:LinkHeadersFooters
            Document dstDoc = new Document(JoiningAppendingDir + "TestFile.Destination.doc");
            Document srcDoc = new Document(JoiningAppendingDir + "TestFile.Source.doc");

            // Set the appended document to appear on a new page
            srcDoc.FirstSection.PageSetup.SectionStart = SectionStart.NewPage;

            // Link the headers and footers in the source document to the previous section
            // This will override any headers or footers already found in the source document
            srcDoc.FirstSection.HeadersFooters.LinkToPrevious(true);

            dstDoc.AppendDocument(srcDoc, ImportFormatMode.KeepSourceFormatting);
            
            dstDoc.Save(ArtifactsDir + "LinkHeadersFooters.docx");
            //ExEnd:LinkHeadersFooters
        }
    }
}