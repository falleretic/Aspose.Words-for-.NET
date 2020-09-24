using Aspose.Words.Fonts;
using NUnit.Framework;

namespace Aspose.Words.Examples.CSharp
{
    class WorkingWithFontSettings : TestDataHelper
    {
        [Test]
        public static void FontSettingsWithLoadOption()
        {
            // ExStart:FontSettingsWithLoadOption
            FontSettings fontSettings = new FontSettings();
            // init font settings
            Words.LoadOptions loadOptions = new Words.LoadOptions();
            loadOptions.FontSettings = fontSettings;
            Document doc1 = new Document(RenderingPrintingDir + "Rendering.docx", loadOptions);

            Words.LoadOptions loadOptions2 = new Words.LoadOptions();
            loadOptions2.FontSettings = fontSettings;
            Document doc2 = new Document(RenderingPrintingDir + "Rendering.docx", loadOptions2);
            // ExEnd:FontSettingsWithLoadOption   
        }

        [Test]
        public static void FontSettingsDefaultInstance()
        {
            // ExStart:FontSettingsFontSource
            // ExStart:FontSettingsDefaultInstance
            FontSettings fontSettings = FontSettings.DefaultInstance;
            // ExEnd:FontSettingsDefaultInstance   
            fontSettings.SetFontsSources(new FontSourceBase[]
             {
                 new SystemFontSource(),
                 new FolderFontSource("C:\\MyFonts\\", true)
             });
            // ExEnd:FontSettingsFontSource

            // init font settings
            Words.LoadOptions loadOptions = new Words.LoadOptions();
            loadOptions.FontSettings = fontSettings;
            Document doc1 = new Document(RenderingPrintingDir + "Rendering.docx", loadOptions);

            Words.LoadOptions loadOptions2 = new Words.LoadOptions();
            loadOptions2.FontSettings = fontSettings;
            Document doc2 = new Document(RenderingPrintingDir + "Rendering.docx", loadOptions2);
        }
    }
}
