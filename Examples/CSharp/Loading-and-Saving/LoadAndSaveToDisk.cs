using NUnit.Framework;

namespace Aspose.Words.Examples.CSharp.Loading_Saving
{
    class LoadAndSaveToDisk : TestDataHelper
    {
        [Test]
        public static void Run()
        {
            //ExStart:LoadAndSave
            //ExStart:OpenDocument
            Document doc = new Document(QuickStartDir + "Document.doc");
            //ExEnd:OpenDocument
            doc.Save(ArtifactsDir + "LoadAndSaveToDisk.docx");
            //ExEnd:LoadAndSave
        }
    }
}