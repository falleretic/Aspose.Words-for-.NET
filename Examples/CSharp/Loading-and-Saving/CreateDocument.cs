using NUnit.Framework;

namespace Aspose.Words.Examples.CSharp.Loading_Saving
{
    class CreateDocument : TestDataHelper
    {
        [Test]
        public static void Run()
        {
            //ExStart:CreateDocument
            Document doc = new Document();

            // Use a document builder to add content to the document.
            DocumentBuilder builder = new DocumentBuilder(doc);
            builder.Writeln("Hello World!");

            doc.Save(ArtifactsDir + "CreateDocument.docx");
            //ExEnd:CreateDocument
        }
    }
}