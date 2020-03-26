using System;

namespace Aspose.Words.Examples.CSharp.Programming_Documents.Working_with_Fields
{
    class InsertFormFields
    {
        public static void Run()
        {
            // ExStart:InsertFormFields
            // The path to the documents directory.
            string dataDir = RunExamples.GetDataDir_WorkingWithFields();

            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            string[] items = { "One", "Two", "Three" };
            builder.InsertComboBox("DropDown", items, 0);
            // ExEnd:InsertFormFields
            Console.WriteLine("\nForm fields inserted successfully.");
        }
    }
}