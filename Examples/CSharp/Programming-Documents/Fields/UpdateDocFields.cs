using NUnit.Framework;

namespace Aspose.Words.Examples.CSharp.Programming_Documents.Working_with_Fields
{
    class UpdateDocFields : TestDataHelper
    {
        [Test]
        public static void Run()
        {
            //ExStart:UpdateDocFields
            Document doc = new Document(FieldsDir + "Rendering.doc");
            // This updates all fields in the document
            doc.UpdateFields();
            
            doc.Save(ArtifactsDir + "Rendering.UpdateFields.pdf");
            //ExEnd:UpdateDocFields
        }
    }
}