using Aspose.Words.Fields;
using NUnit.Framework;

namespace Aspose.Words.Examples.CSharp.Programming_Documents.Working_with_Fields
{
    class ConvertFieldsInBody : TestDataHelper
    {
        [Test]
        public static void Run()
        {
            //ExStart:ConvertFieldsInBody
            Document doc = new Document(FieldsDir + "TestFile.doc");

            // Pass the appropriate parameters to convert PAGE fields encountered to static text only in the body of the first section
            FieldsHelper.ConvertFieldsToStaticText(doc.FirstSection.Body, FieldType.FieldPage);

            // Save the document with fields transformed to disk
            doc.Save(ArtifactsDir + "TestFile.doc");
            //ExEnd:ConvertFieldsInBody
        }
    }
}