namespace Aspose.Words.Examples.CSharp.Programming_Documents.Joining_and_Appending
{
    class SimpleAppendDocument : TestDataHelper
    {
        public static void Run()
        {
            Document dstDoc = new Document(JoiningAppendingDir + "TestFile.Destination.doc");
            Document srcDoc = new Document(JoiningAppendingDir + "TestFile.Source.doc");

            // Append the source document to the destination document using no extra options
            dstDoc.AppendDocument(srcDoc, ImportFormatMode.KeepSourceFormatting);

            dstDoc.Save(ArtifactsDir + "SimpleAppendDocument.docx");
        }
    }
}