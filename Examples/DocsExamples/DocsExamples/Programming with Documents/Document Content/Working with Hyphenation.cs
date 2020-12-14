using System.IO;
using Aspose.Words;
using NUnit.Framework;

namespace DocsExamples.Programming_with_Documents.Document_Content
{
    internal class WorkingWithHyphenation : DocsExamplesBase
    {
        [Test]
        public static void HyphenateWordsOfLanguages()
        {
            //ExStart:HyphenateWordsOfLanguages
            Document doc = new Document(MyDir + "German text.docx");

            Hyphenation.RegisterDictionary("en-US", MyDir + "hyph_en_US.dic");
            Hyphenation.RegisterDictionary("de-CH", MyDir + "hyph_de_CH.dic");

            doc.Save(ArtifactsDir + "WorkingWithHyphenation.HyphenateWordsOfLanguages.pdf");
            //ExEnd:HyphenateWordsOfLanguages
        }

        [Test]
        public static void LoadHyphenationDictionaryForLanguage()
        {
            //ExStart:LoadHyphenationDictionaryForLanguage
            Document doc = new Document(MyDir + "German text.docx");
            
            Stream stream = File.OpenRead(MyDir + "hyph_de_CH.dic");
            Hyphenation.RegisterDictionary("de-CH", stream);

            doc.Save(ArtifactsDir + "WorkingWithHyphenation.LoadHyphenationDictionaryForLanguage.pdf");
            //ExEnd:LoadHyphenationDictionaryForLanguage
        }
    }
}