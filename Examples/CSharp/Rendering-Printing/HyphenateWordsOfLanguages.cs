using NUnit.Framework;

namespace Aspose.Words.Examples.CSharp.Rendering_and_Printing
{
    class HyphenateWordsOfLanguages : TestDataHelper
    {
        [Test]
        public static void Run()
        {
            //ExStart:HyphenateWordsOfLanguages
            // Load the documents which store the shapes we want to render
            Document doc = new Document(MailMergeDir + "TestFile RenderShape.doc");

            Hyphenation.RegisterDictionary("en-US", MailMergeDir + "hyph_en_US.dic");
            Hyphenation.RegisterDictionary("de-CH", MailMergeDir + "hyph_de_CH.dic");

            doc.Save(ArtifactsDir + "HyphenateWordsOfLanguages.pdf");
            //ExEnd:HyphenateWordsOfLanguages
        }
    }
}