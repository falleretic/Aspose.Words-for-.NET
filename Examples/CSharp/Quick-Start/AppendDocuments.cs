using NUnit.Framework;

namespace Aspose.Words.Examples.CSharp
{
    class AppendDocuments : TestDataHelper
    {
        [Test]
        public static void Run()
        {
            // Load the destination and source documents from disk
            Document dstDoc = new Document(QuickStartDir + "Document insertion destination.docx");
            Document srcDoc = new Document(QuickStartDir + "Document.docx");

            // Append the source document to the destination document while keeping the original formatting of the source document
            dstDoc.AppendDocument(srcDoc, ImportFormatMode.KeepSourceFormatting);
            
            dstDoc.Save(ArtifactsDir + "AppendDocuments.docx");
        }
    }
}