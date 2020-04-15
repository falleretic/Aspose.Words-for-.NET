using Aspose.Words.Fields;
using NUnit.Framework;

namespace Aspose.Words.Examples.CSharp.Programming_Documents.Working_with_Fields
{
    class RemoveField : TestDataHelper
    {
        [Test]
        public static void Run()
        {
            //ExStart:RemoveField
            Document doc = new Document(FieldsDir + "Field.RemoveField.doc");
            
            Field field = doc.Range.Fields[0];
            // Calling this method completely removes the field from the document
            field.Remove();
            //ExEnd:RemoveField
        }
    }
}