using NUnit.Framework;

namespace Aspose.Words.Examples.CSharp.DocumentEx
{
    class CloningDocument : TestDataHelper
    {
        [Test]
        public static void Run()
        {
            //ExStart:CloningDocument
            Document doc = new Document(DocumentDir + "Document.docx");

            Document clone = doc.Clone();
            clone.Save(ArtifactsDir + "CloningDocument.doc");
            //ExEnd:CloningDocument
        }
    }
}