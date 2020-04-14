using System.IO;

namespace Aspose.Words.Examples.CSharp.Rendering_and_Printing
{
    class LoadHyphenationDictionaryForLanguage : TestDataHelper
    {
        public static void Run()
        {
            //ExStart:LoadHyphenationDictionaryForLanguage
            // Load the documents which store the shapes we want to render
            Document doc = new Document(MailMergeDir + "TestFile RenderShape.doc");
            
            Stream stream = File.OpenRead(MailMergeDir + "hyph_de_CH.dic");
            Hyphenation.RegisterDictionary("de-CH", stream);

            doc.Save(ArtifactsDir + "LoadHyphenationDictionaryForLanguage.pdf");
            //ExEnd:LoadHyphenationDictionaryForLanguage
        }
    }
}