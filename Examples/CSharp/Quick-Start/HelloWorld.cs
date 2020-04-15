using NUnit.Framework;

namespace Aspose.Words.Examples.CSharp.Quick_Start
{
    class HelloWorld : TestDataHelper
    {
        [Test]
        public static void Run()
        {
            Document doc = new Document();

            // DocumentBuilder provides members to easily add content to a document
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Write a new paragraph in the document with the text "Hello World!"
            builder.Writeln("Hello World!");

            // Save the document in DOCX format. The format to save as is inferred from the extension of the file name
            // Aspose.Words supports saving any document in many more formats
            doc.Save(ArtifactsDir + "HelloWorld.docx");
        }
    }
}