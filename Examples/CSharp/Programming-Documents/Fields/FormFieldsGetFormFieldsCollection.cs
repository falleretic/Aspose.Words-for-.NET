using System;
using Aspose.Words.Fields;

namespace Aspose.Words.Examples.CSharp.Programming_Documents.Working_with_Fields
{
    class FormFieldsGetFormFieldsCollection : TestDataHelper
    {
        public static void Run()
        {
            //ExStart:FormFieldsGetFormFieldsCollection
            Document doc = new Document(FieldsDir + "FormFields.doc");
            FormFieldCollection formFields = doc.Range.FormFields;
            //ExEnd:FormFieldsGetFormFieldsCollection
        }
    }
}