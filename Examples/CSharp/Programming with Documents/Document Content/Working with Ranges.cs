using NUnit.Framework;

namespace Aspose.Words.Examples.CSharp.Ranges
{
    class WorkingWithRanges : TestDataHelper
    {
        [Test]
        public static void RangesDeleteText()
        {
            //ExStart:RangesDeleteText
            Document doc = new Document(RangeDir + "Document.docx");
            doc.Sections[0].Range.Delete();
            //ExEnd:RangesDeleteText
        }

        [Test]
        public static void RangesGetText()
        {
            //ExStart:RangesGetText
            Document doc = new Document(RangeDir + "Document.docx");
            string text = doc.Range.Text;
            //ExEnd:RangesGetText
        }
    }
}