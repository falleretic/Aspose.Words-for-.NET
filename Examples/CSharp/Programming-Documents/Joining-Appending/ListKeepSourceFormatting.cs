namespace Aspose.Words.Examples.CSharp.Programming_Documents.Joining_and_Appending
{
    class ListKeepSourceFormatting : TestDataHelper
    {
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