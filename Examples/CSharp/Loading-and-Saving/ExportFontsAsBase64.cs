using Aspose.Words.Saving;
using NUnit.Framework;

namespace Aspose.Words.Examples.CSharp.Loading_Saving
{
    class ExportFontsAsBase64 : TestDataHelper
    {
        [Test]
        public static void Run()
        {
            //ExStart:ExportFontsAsBase64
            Document doc = new Document(LoadingSavingDir + "Document.doc");
            
            HtmlSaveOptions saveOptions = new HtmlSaveOptions();
            saveOptions.ExportFontResources = true;
            saveOptions.ExportFontsAsBase64 = true;
            
            doc.Save(ArtifactsDir + "ExportFontsAsBase64.html", saveOptions);
            //ExEnd:ExportFontsAsBase64
        }
    }
}