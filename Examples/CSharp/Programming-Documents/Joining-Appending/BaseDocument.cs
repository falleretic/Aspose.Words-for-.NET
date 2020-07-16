using NUnit.Framework;

namespace Aspose.Words.Examples.CSharp
{
    class BaseDocument : TestDataHelper
    {
        [Test]
        public static void Run()
        {
            //ExStart:BaseDocument
            Document dstDoc = new Document();
            Document srcDoc = new Document(JoiningAppendingDir + "TestFile.Source.doc");

            // The destination document is not actually empty which often causes a blank page to appear before the appended document
            // This is due to the base document having an empty section and the new document being started on the next page
            // Remove all content from the destination document before appending
            dstDoc.RemoveAllChildren();
            dstDoc.AppendDocument(srcDoc, ImportFormatMode.KeepSourceFormatting);
            
            dstDoc.Save(ArtifactsDir + "BaseDocument.docx");
            //ExEnd:BaseDocument
        }
    }
}