using NUnit.Framework;

namespace Aspose.Words.Examples.CSharp.Programming_Documents.Working_with_Ranges
{
    class RangesDeleteText : TestDataHelper
    {
        [Test]
        public static void Run()
        {
            //ExStart:RangesDeleteText
            Document doc = new Document(RangeDir + "Document.doc");
            doc.Sections[0].Range.Delete();
            //ExEnd:RangesDeleteText
        }
    }
}