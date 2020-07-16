using NUnit.Framework;

namespace Aspose.Words.Examples.CSharp.DocumentEx
{
    class AccessStyles : TestDataHelper
    {
        [Test]
        public static void Run()
        {
            //ExStart:AccessStyles
            Document doc = new Document(DocumentDir + "TestFile.doc");

            // Get styles collection from document
            StyleCollection styles = doc.Styles;
            string styleName = "";

            // Iterate through all the styles
            foreach (Style style in styles)
            {
                if (styleName == "")
                {
                    styleName = style.Name;
                }
                else
                {
                    styleName = styleName + ", " + style.Name;
                }
            }
            //ExEnd:AccessStyles
        }
    }
}