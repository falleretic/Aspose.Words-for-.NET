using System.IO;
using NUnit.Framework;

namespace Aspose.Words.Examples.CSharp
{
    class LoadHyphenationDictionaryForLanguage : TestDataHelper
    {
        [Test]
        public static void Run()
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