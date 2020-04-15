using Aspose.Words.Fields;
using NUnit.Framework;

namespace Aspose.Words.Examples.CSharp.Programming_Documents.Working_with_Fields
{
    class FormFieldsWorkWithProperties : TestDataHelper
    {
        [Test]
        public static void Run()
        {
            //ExStart:FormFieldsWorkWithProperties
            Document doc = new Document(FieldsDir + "FormFields.doc");
            FormField formField = doc.Range.FormFields[3];

            if (formField.Type.Equals(FieldType.FieldFormTextInput))
                formField.Result = "My name is " + formField.Name;
            //ExEnd:FormFieldsWorkWithProperties            
        }
    }
}