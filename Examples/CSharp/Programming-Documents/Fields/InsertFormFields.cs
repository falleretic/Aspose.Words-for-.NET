using NUnit.Framework;

namespace Aspose.Words.Examples.CSharp.Programming_Documents.Working_with_Fields
{
    class InsertFormFields : TestDataHelper
    {
        [Test]
        public static void Run()
        {
            //ExStart:InsertFormFields
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            string[] items = { "One", "Two", "Three" };
            builder.InsertComboBox("DropDown", items, 0);
            //ExEnd:InsertFormFields
        }
    }
}