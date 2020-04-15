using NUnit.Framework;

namespace Aspose.Words.Examples.CSharp.Programming_Documents.Joining_and_Appending
{
    class UnlinkHeadersFooters : TestDataHelper
    {
        [Test]
        public static void Run()
        {
            //ExStart:UnlinkHeadersFooters
            Document dstDoc = new Document(JoiningAppendingDir + "TestFile.Destination.doc");
            Document srcDoc = new Document(JoiningAppendingDir + "TestFile.Source.doc");

            // Unlink the headers and footers in the source document to stop this from continuing the headers and footers
            // From the destination document
            srcDoc.FirstSection.HeadersFooters.LinkToPrevious(false);

            dstDoc.AppendDocument(srcDoc, ImportFormatMode.KeepSourceFormatting);
            
            dstDoc.Save(ArtifactsDir + "UnlinkHeadersFooters.docx");
            //ExEnd:UnlinkHeadersFooters
        }
    }
}