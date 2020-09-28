using System.IO;
using NUnit.Framework;

namespace Aspose.Words.Examples.CSharp
{
    class WorkingWithHyphenation : TestDataHelper
    {
        [Test]
        public static void HyphenateWordsOfLanguages()
        {
            //ExStart:HyphenateWordsOfLanguages
            // Load the documents which store the shapes we want to render
            Document doc = new Document(RenderingPrintingDir + "German text.docx");

            Hyphenation.RegisterDictionary("en-US", RenderingPrintingDir + "hyph_en_US.dic");
            Hyphenation.RegisterDictionary("de-CH", RenderingPrintingDir + "hyph_de_CH.dic");

            doc.Save(ArtifactsDir + "Hyphenation.Dictionary.Registered.pdf");
            //ExEnd:HyphenateWordsOfLanguages
        }

        [Test]
        public static void LoadHyphenationDictionaryForLanguage()
        {
            //ExStart:LoadHyphenationDictionaryForLanguage
            // Load the documents which store the shapes we want to render
            Document doc = new Document(RenderingPrintingDir + "German text.docx");
            
            Stream stream = File.OpenRead(RenderingPrintingDir + "hyph_de_CH.dic");
            Hyphenation.RegisterDictionary("de-CH", stream);

            doc.Save(ArtifactsDir + "Hyphenation.Stream.pdf");
            //ExEnd:LoadHyphenationDictionaryForLanguage
        }
    }
}