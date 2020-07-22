using System.Linq;
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
            Document doc = new Document(FieldsDir + "Linked fields.docx");

            // Pass the appropriate parameters to convert all IF fields encountered in the document (including headers and footers) to static text
            doc.Range.Fields.Where(f => f.Type == FieldType.FieldIf).ToList().ForEach(f => f.Unlink());

            // Save the document with fields transformed to disk
            doc.Save(ArtifactsDir + "TestFile.doc");
            //ExEnd:ConvertFieldsInDocument
        }
    }
}