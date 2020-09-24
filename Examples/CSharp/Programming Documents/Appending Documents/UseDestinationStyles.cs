using NUnit.Framework;

namespace Aspose.Words.Examples.CSharp
{
    class UseDestinationStyles : TestDataHelper
    {
        [Test]
        public static void Run()
        {
            //ExStart:UseDestinationStyles
            Document srcDoc = new Document(JoiningAppendingDir + "Document source.docx");
            Document dstDoc = new Document(JoiningAppendingDir + "Northwind traders.docx");

            // Append the source document using the styles of the destination document
            dstDoc.AppendDocument(srcDoc, ImportFormatMode.UseDestinationStyles);

            dstDoc.Save(ArtifactsDir + "UseDestinationStyles.docx");
            //ExEnd:UseDestinationStyles
        }
    }
}