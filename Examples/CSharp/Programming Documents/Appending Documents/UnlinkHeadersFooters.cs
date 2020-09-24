using NUnit.Framework;

namespace Aspose.Words.Examples.CSharp
{
    class UnlinkHeadersFooters : TestDataHelper
    {
        [Test]
        public static void Run()
        {
            //ExStart:UnlinkHeadersFooters
            Document srcDoc = new Document(JoiningAppendingDir + "Document source.docx");
            Document dstDoc = new Document(JoiningAppendingDir + "Northwind traders.docx");

            // Unlink the headers and footers in the source document to stop this from continuing the headers and footers
            // From the destination document
            srcDoc.FirstSection.HeadersFooters.LinkToPrevious(false);

            dstDoc.AppendDocument(srcDoc, ImportFormatMode.KeepSourceFormatting);
            
            dstDoc.Save(ArtifactsDir + "UnlinkHeadersFooters.docx");
            //ExEnd:UnlinkHeadersFooters
        }
    }
}