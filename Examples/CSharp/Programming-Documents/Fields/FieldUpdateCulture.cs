using Aspose.Words.Fields;
using NUnit.Framework;

namespace Aspose.Words.Examples.CSharp.Fields
{
    class FieldUpdateCulture : TestDataHelper
    {
        [Test]
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