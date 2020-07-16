using NUnit.Framework;

namespace Aspose.Words.Examples.CSharp.Styles
{
    class ChangeStyleOfTocLevel
    {
        [Test]
        public static void Run()
        {
            //ExStart:ChangeStyleOfTOCLevel
            Document doc = new Document();
            // Retrieve the style used for the first level of the TOC and change the formatting of the style
            doc.Styles[StyleIdentifier.Toc1].Font.Bold = true;
            //ExEnd:ChangeStyleOfTOCLevel
        }
    }
}