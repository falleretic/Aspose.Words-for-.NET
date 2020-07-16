using Aspose.Words.Fields;
using NUnit.Framework;

namespace Aspose.Words.Examples.CSharp.Fields
{
    class ConvertFieldsInParagraph : TestDataHelper
    {
        [Test]
        public static void Run()
        {
            //ExStart:ConvertFieldsInParagraph
            Document doc = new Document(FieldsDir + "TestFile.doc");

            // Pass the appropriate parameters to convert all IF fields to static text that are encountered only in the last 
            // paragraph of the document
            FieldsHelper.ConvertFieldsToStaticText(doc.FirstSection.Body.LastParagraph, FieldType.FieldIf);

            // Save the document with fields transformed to disk
            doc.Save(ArtifactsDir + "TestFile.doc");
            //ExEnd:ConvertFieldsInParagraph
        }
    }
}