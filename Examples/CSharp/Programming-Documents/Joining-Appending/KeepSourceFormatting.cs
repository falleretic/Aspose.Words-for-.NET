using NUnit.Framework;

namespace Aspose.Words.Examples.CSharp
{
    class KeepSourceFormatting : TestDataHelper
    {
        [Test]
        public static void Run()
        {
            //ExStart:KeepSourceFormatting
            Document dstDoc = new Document(JoiningAppendingDir + "TestFile.Destination.doc");
            Document srcDoc = new Document(JoiningAppendingDir + "TestFile.Source.doc");

            // Keep the formatting from the source document when appending it to the destination document
            dstDoc.AppendDocument(srcDoc, ImportFormatMode.KeepSourceFormatting);

            // Save the joined document to disk
            dstDoc.Save(ArtifactsDir + "KeepSourceFormatting.docx");
            //ExEnd:KeepSourceFormatting
        }
    }
}