using Aspose.Words.Fields;
using NUnit.Framework;

namespace Aspose.Words.Examples.CSharp.Fields
{
    class ConvertFieldsInDocument : TestDataHelper
    {
        [Test]
        public static void Run()
        {
            //ExStart:ConvertFieldsInDocument
            Document doc = new Document(FieldsDir + "TestFile.doc");

            // Pass the appropriate parameters to convert all IF fields encountered in the document (including headers and footers) to static text
            FieldsHelper.ConvertFieldsToStaticText(doc, FieldType.FieldIf);

            // Save the document with fields transformed to disk
            doc.Save(ArtifactsDir + "TestFile.doc");
            //ExEnd:ConvertFieldsInDocument
        }
    }
}