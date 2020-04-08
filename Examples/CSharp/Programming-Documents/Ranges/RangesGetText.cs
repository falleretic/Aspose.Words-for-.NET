namespace Aspose.Words.Examples.CSharp.Programming_Documents.Working_with_Ranges
{
    class RangesGetText : TestDataHelper
    {
        public static void Run()
        {
            //ExStart:RangesGetText
            Document doc = new Document(RangeDir + "Document.doc");
            string text = doc.Range.Text;
            //ExEnd:RangesGetText
        }
    }
}