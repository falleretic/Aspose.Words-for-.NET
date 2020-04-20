using Aspose.Words.Fields;
using System;
using NUnit.Framework;

namespace Aspose.Words.Examples.CSharp.Programming_Documents.Working_with_Fields
{
    class FieldDisplayResults : TestDataHelper
    {
        [Test]
        public static void Run()
        {
            //ExStart:FieldDisplayResults
            Document document = new Document(LoadingSavingDir + "Document.docx");
            document.UpdateFields();

            foreach (Field field in document.Range.Fields)
                Console.WriteLine(field.DisplayResult);
            //ExEnd:FieldDisplayResults
        }
    }
}