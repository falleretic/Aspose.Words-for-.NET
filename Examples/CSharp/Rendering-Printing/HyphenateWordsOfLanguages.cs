using NUnit.Framework;

namespace Aspose.Words.Examples.CSharp
{
    class HyphenateWordsOfLanguages : TestDataHelper
    {
        [Test]
        public static void Run()
        {
            //ExStart:HyphenateWordsOfLanguages
            // Load the documents which store the shapes we want to render
            Document doc = new Document(RenderingPrintingDir + "TestFile RenderShape.doc");

            Hyphenation.RegisterDictionary("en-US", RenderingPrintingDir + "hyph_en_US.dic");
            Hyphenation.RegisterDictionary("de-CH", RenderingPrintingDir + "hyph_de_CH.dic");

            doc.Save(ArtifactsDir + "HyphenateWordsOfLanguages.pdf");
            //ExEnd:HyphenateWordsOfLanguages
        }
    }
}