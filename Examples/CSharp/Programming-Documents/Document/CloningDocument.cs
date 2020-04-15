using NUnit.Framework;

namespace Aspose.Words.Examples.CSharp.Programming_Documents.Working_With_Document
{
    class CloningDocument : TestDataHelper
    {
        [Test]
        public static void Run()
        {
            //ExStart:CloningDocument
            Document doc = new Document(DocumentDir + "TestFile.doc");

            Document clone = doc.Clone();
            clone.Save(ArtifactsDir + "TestFile_clone.doc");
            //ExEnd:CloningDocument
        }
    }
}