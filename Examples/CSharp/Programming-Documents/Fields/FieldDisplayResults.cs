using Aspose.Words.Fields;
using System;

namespace Aspose.Words.Examples.CSharp.Programming_Documents.Working_with_Fields
{
    class FieldDisplayResults : TestDataHelper
    {
        public static void Run()
        {
            //ExStart:FieldDisplayResults
            Document document = new Document(FieldsDir + "Document.docx");
            document.UpdateFields();

            foreach (Field field in document.Range.Fields)
                Console.WriteLine(field.DisplayResult);
            //ExEnd:FieldDisplayResults
        }
    }
}