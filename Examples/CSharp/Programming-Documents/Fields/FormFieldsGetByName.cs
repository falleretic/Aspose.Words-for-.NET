using Aspose.Words.Fields;
using NUnit.Framework;

namespace Aspose.Words.Examples.CSharp.Fields
{
    class FormFieldsGetByName : TestDataHelper
    {
        [Test]
        public static void Run()
        {
            //ExStart:FormFieldsGetByName
            Document doc = new Document(FieldsDir + "FormFields.doc");
            FormFieldCollection documentFormFields = doc.Range.FormFields;

            FormField formField1 = documentFormFields[3];
            FormField formField2 = documentFormFields["Text2"];
            //ExEnd:FormFieldsGetByName
        }
    }
}