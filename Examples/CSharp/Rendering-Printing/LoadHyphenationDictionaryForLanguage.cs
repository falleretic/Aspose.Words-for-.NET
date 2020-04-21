using System.IO;
using NUnit.Framework;

namespace Aspose.Words.Examples.CSharp.Rendering_and_Printing
{
    class LoadHyphenationDictionaryForLanguage : TestDataHelper
    {
        [Test]
        public static void Run()
        {
            //ExStart:LoadHyphenationDictionaryForLanguage
            // Load the documents which store the shapes we want to render
            Document doc = new Document(RenderingPrintingDir + "TestFile RenderShape.doc");
            
            Stream stream = File.OpenRead(RenderingPrintingDir + "hyph_de_CH.dic");
            Hyphenation.RegisterDictionary("de-CH", stream);

            doc.Save(ArtifactsDir + "LoadHyphenationDictionaryForLanguage.pdf");
            //ExEnd:LoadHyphenationDictionaryForLanguage
        }
    }
}