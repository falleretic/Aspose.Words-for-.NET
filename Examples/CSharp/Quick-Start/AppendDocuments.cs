using NUnit.Framework;

namespace Aspose.Words.Examples.CSharp.Quick_Start
{
    class AppendDocuments : TestDataHelper
    {
        [Test]
        public static void Run()
        {
            // Load the destination and source documents from disk
            Document dstDoc = new Document(QuickStartDir + "TestFile.Destination.doc");
            Document srcDoc = new Document(QuickStartDir + "TestFile.Source.doc");

            // Append the source document to the destination document while keeping the original formatting of the source document
            dstDoc.AppendDocument(srcDoc, ImportFormatMode.KeepSourceFormatting);
            
            dstDoc.Save(ArtifactsDir + "TestFile.Destination.doc");
        }
    }
}