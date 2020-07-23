using NUnit.Framework;

namespace Aspose.Words.Examples.CSharp.Ranges
{
    class RangesDeleteText : TestDataHelper
    {
        [Test]
        public static void Run()
        {
            //ExStart:RangesDeleteText
            Document doc = new Document(RangeDir + "Document.docx");
            doc.Sections[0].Range.Delete();
            //ExEnd:RangesDeleteText
        }
    }
}