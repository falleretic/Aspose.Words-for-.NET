using Aspose.Words.Fields;

namespace Aspose.Words.Examples.CSharp.Programming_Documents.Working_with_Fields
{
    class FieldUpdateCulture : TestDataHelper
    {
        public static void Run()
        {
            //ExStart:FieldUpdateCultureProvider
            Document doc = new Document(FieldsDir + "FieldUpdateCultureProvider.docx");

            doc.FieldOptions.FieldUpdateCultureSource = FieldUpdateCultureSource.FieldCode;
            doc.FieldOptions.FieldUpdateCultureProvider = new FieldUpdateCultureProvider();

            doc.Save(ArtifactsDir + "Field.FieldUpdateCultureProvider.pdf");
            //ExEnd:FieldUpdateCultureProvider
        }
    }
}