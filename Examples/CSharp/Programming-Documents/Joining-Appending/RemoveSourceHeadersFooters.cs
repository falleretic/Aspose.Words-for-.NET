using NUnit.Framework;

namespace Aspose.Words.Examples.CSharp
{
    class RemoveSourceHeadersFooters : TestDataHelper
    {
        [Test]
        public static void Run()
        {
            //ExStart:RemoveSourceHeadersFooters
            Document dstDoc = new Document(JoiningAppendingDir + "TestFile.Destination.doc");
            Document srcDoc = new Document(JoiningAppendingDir + "TestFile.Source.doc");

            // Remove the headers and footers from each of the sections in the source document
            foreach (Section section in srcDoc.Sections)
            {
                section.ClearHeadersFooters();
            }

            // Even after the headers and footers are cleared from the source document, the "LinkToPrevious" setting 
            // For HeadersFooters can still be set. This will cause the headers and footers to continue from the destination 
            // Document. This should set to false to avoid this behavior
            srcDoc.FirstSection.HeadersFooters.LinkToPrevious(false);

            dstDoc.AppendDocument(srcDoc, ImportFormatMode.KeepSourceFormatting);
            
            dstDoc.Save(ArtifactsDir + "RemoveSourceHeadersFooters.docx");
            //ExEnd:RemoveSourceHeadersFooters
        }
    }
}