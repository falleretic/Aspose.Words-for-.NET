using System.Linq;
using Aspose.Words.Fields;
using NUnit.Framework;

namespace Aspose.Words.Examples.CSharp.Fields
{
    class ConvertFieldsInBody : TestDataHelper
    {
        [Test]
        public static void Run()
        {
            //ExStart:ConvertFieldsInBody
            Document doc = new Document(FieldsDir + "Linked fields.docx");

            // Pass the appropriate parameters to convert PAGE fields encountered to static text only in the body of the first section
            doc.FirstSection.Body.Range.Fields.Where(f => f.Type == FieldType.FieldPage).ToList().ForEach(f => f.Unlink());

            // Save the document with fields transformed to disk
            doc.Save(ArtifactsDir + "TestFile.doc");
            //ExEnd:ConvertFieldsInBody
        }
    }
}