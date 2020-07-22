using Aspose.Words.Fields;
using NUnit.Framework;

namespace Aspose.Words.Examples.CSharp.Fields
{
    class RemoveField : TestDataHelper
    {
        [Test]
        public static void Run()
        {
            //ExStart:RemoveField
            Document doc = new Document(FieldsDir + "Various fields.docx");
            
            Field field = doc.Range.Fields[0];
            // Calling this method completely removes the field from the document
            field.Remove();
            //ExEnd:RemoveField
        }
    }
}