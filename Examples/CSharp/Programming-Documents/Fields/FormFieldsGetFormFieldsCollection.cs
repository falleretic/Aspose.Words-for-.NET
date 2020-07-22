using Aspose.Words.Fields;
using NUnit.Framework;

namespace Aspose.Words.Examples.CSharp.Fields
{
    class FormFieldsGetFormFieldsCollection : TestDataHelper
    {
        [Test]
        public static void Run()
        {
            //ExStart:FormFieldsGetFormFieldsCollection
            Document doc = new Document(FieldsDir + "Form fields.docx");
            FormFieldCollection formFields = doc.Range.FormFields;
            //ExEnd:FormFieldsGetFormFieldsCollection
        }
    }
}