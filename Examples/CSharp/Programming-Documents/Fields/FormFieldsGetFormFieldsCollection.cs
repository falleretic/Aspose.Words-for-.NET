using Aspose.Words.Fields;
using NUnit.Framework;

namespace Aspose.Words.Examples.CSharp.Programming_Documents.Working_with_Fields
{
    class FormFieldsGetFormFieldsCollection : TestDataHelper
    {
        [Test]
        public static void Run()
        {
            //ExStart:FormFieldsGetFormFieldsCollection
            Document doc = new Document(FieldsDir + "FormFields.doc");
            FormFieldCollection formFields = doc.Range.FormFields;
            //ExEnd:FormFieldsGetFormFieldsCollection
        }
    }
}