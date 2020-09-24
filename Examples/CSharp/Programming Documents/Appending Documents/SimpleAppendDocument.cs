using NUnit.Framework;

namespace Aspose.Words.Examples.CSharp
{
    class SimpleAppendDocument : TestDataHelper
    {
        [Test]
        public static void Run()
        {
            Document srcDoc = new Document(JoiningAppendingDir + "Document source.docx");
            Document dstDoc = new Document(JoiningAppendingDir + "Northwind traders.docx");

            // Append the source document to the destination document using no extra options
            dstDoc.AppendDocument(srcDoc, ImportFormatMode.KeepSourceFormatting);

            dstDoc.Save(ArtifactsDir + "SimpleAppendDocument.docx");
        }
    }
}