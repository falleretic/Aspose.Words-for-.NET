using Aspose.Words;
using NUnit.Framework;

namespace SiteExamples.Programming_with_Documents.Document_Content
{
    class WorkingWithRanges : SiteExamplesBase
    {
        [Test]
        public static void RangesDeleteText()
        {
            //ExStart:RangesDeleteText
            Document doc = new Document(MyDir + "Document.docx");
            doc.Sections[0].Range.Delete();
            //ExEnd:RangesDeleteText
        }

        [Test]
        public static void RangesGetText()
        {
            //ExStart:RangesGetText
            Document doc = new Document(MyDir + "Document.docx");
            string text = doc.Range.Text;
            //ExEnd:RangesGetText
        }
    }
}