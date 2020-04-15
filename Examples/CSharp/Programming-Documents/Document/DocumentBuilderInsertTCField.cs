using NUnit.Framework;

namespace Aspose.Words.Examples.CSharp.Programming_Documents.Working_With_Document
{
    class DocumentBuilderInsertTCField : TestDataHelper
    {
        [Test]
        public static void Run()
        {
            //ExStart:DocumentBuilderInsertTCField
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Insert a TC field at the current document builder position
            builder.InsertField("TC \"Entry Text\" \\f t");

            doc.Save(ArtifactsDir + "DocumentBuilderInsertTCField.doc");
            //ExEnd:DocumentBuilderInsertTCField
        }
    }
}