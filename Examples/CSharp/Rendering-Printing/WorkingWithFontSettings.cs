using Aspose.Words.Fonts;
using NUnit.Framework;

namespace Aspose.Words.Examples.CSharp.Rendering_Printing
{
    class WorkingWithFontSettings : TestDataHelper
    {
        [Test]
        public static void FontSettingsWithLoadOption()
        {
            // ExStart:FontSettingsWithLoadOption
            FontSettings fontSettings = new FontSettings();
            // init font settings
            LoadOptions loadOptions = new LoadOptions();
            loadOptions.FontSettings = fontSettings;
            Document doc1 = new Document(RenderingPrintingDir + "MyDocument.docx", loadOptions);

            LoadOptions loadOptions2 = new LoadOptions();
            loadOptions2.FontSettings = fontSettings;
            Document doc2 = new Document(RenderingPrintingDir + "MyDocument.docx", loadOptions2);
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
                 new FolderFontSource("/home/user/MyFonts", true)
             });
            // ExEnd:FontSettingsFontSource

            // init font settings
            LoadOptions loadOptions = new LoadOptions();
            loadOptions.FontSettings = fontSettings;
            Document doc1 = new Document(RenderingPrintingDir + "MyDocument.docx", loadOptions);

            LoadOptions loadOptions2 = new LoadOptions();
            loadOptions2.FontSettings = fontSettings;
            Document doc2 = new Document(RenderingPrintingDir + "MyDocument.docx", loadOptions2);
        }
    }
}
