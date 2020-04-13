using Aspose.Words.Saving;

namespace Aspose.Words.Examples.CSharp.Loading_Saving
{
    class ExportResourcesUsingHtmlSaveOptions : TestDataHelper
    {
        public static void Run()
        {
            //ExStart:ExportResourcesUsingHtmlSaveOptions
            Document doc = new Document(LoadingSavingDir + "Document.doc");
            
            HtmlSaveOptions saveOptions = new HtmlSaveOptions();
            saveOptions.CssStyleSheetType = CssStyleSheetType.External;
            saveOptions.ExportFontResources = true;
            saveOptions.ResourceFolder = LoadingSavingDir + @"\Resources";
            saveOptions.ResourceFolderAlias = "http://example.com/resources";
            
            doc.Save(ArtifactsDir + "ExportResourcesUsingHtmlSaveOptions.html", saveOptions);
            //ExEnd:ExportResourcesUsingHtmlSaveOptions
        }
    }
}