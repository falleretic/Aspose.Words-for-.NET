using NUnit.Framework;

namespace Aspose.Words.Examples.CSharp.Programming_Documents.Joining_and_Appending
{
    class UseDestinationStyles : TestDataHelper
    {
        [Test]
        public static void Run()
        {
            //ExStart:UseDestinationStyles
            Document dstDoc = new Document(JoiningAppendingDir + "TestFile.Destination.doc");
            Document srcDoc = new Document(JoiningAppendingDir + "TestFile.Source.doc");

            // Append the source document using the styles of the destination document
            dstDoc.AppendDocument(srcDoc, ImportFormatMode.UseDestinationStyles);

            dstDoc.Save(ArtifactsDir + "UseDestinationStyles.docx");
            //ExEnd:UseDestinationStyles
        }
    }
}