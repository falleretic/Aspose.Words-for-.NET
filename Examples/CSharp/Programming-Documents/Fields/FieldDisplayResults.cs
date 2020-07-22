using System;
using Aspose.Words.Fields;
using NUnit.Framework;

namespace Aspose.Words.Examples.CSharp.Fields
{
    class FieldDisplayResults : TestDataHelper
    {
        [Test]
        public static void Run()
        {
            //ExStart:FieldDisplayResults
            //ExStart:UpdateDocFields
            Document document = new Document(FieldsDir + "Various fields.docx");
            // This updates all fields in the document
            document.UpdateFields();
            //ExEnd:UpdateDocFields

            foreach (Field field in document.Range.Fields)
                Console.WriteLine(field.DisplayResult);
            //ExEnd:FieldDisplayResults
        }
    }
}