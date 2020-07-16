using NUnit.Framework;

namespace Aspose.Words.Examples.CSharp
{
    class ListKeepSourceFormatting : TestDataHelper
    {
        [Test]
        public static void Run()
        {
            //ExStart:ListKeepSourceFormatting
            Document dstDoc = new Document(JoiningAppendingDir + "TestFile.DestinationList.doc");
            Document srcDoc = new Document(JoiningAppendingDir + "TestFile.SourceList.doc");

            // Append the content of the document so it flows continuously
            srcDoc.FirstSection.PageSetup.SectionStart = SectionStart.Continuous;

            dstDoc.AppendDocument(srcDoc, ImportFormatMode.KeepSourceFormatting);
            
            dstDoc.Save(ArtifactsDir + "ListKeepSourceFormatting.docx");
            //ExEnd:ListKeepSourceFormatting
        }
    }
}