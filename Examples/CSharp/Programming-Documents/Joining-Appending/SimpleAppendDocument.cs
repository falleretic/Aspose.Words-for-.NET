using NUnit.Framework;

namespace Aspose.Words.Examples.CSharp
{
    class SimpleAppendDocument : TestDataHelper
    {
        [Test]
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