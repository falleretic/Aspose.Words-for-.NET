using Aspose.Words.Saving;
using NUnit.Framework;

namespace Aspose.Words.Examples.CSharp.Loading_Saving
{
    class SaveDocWithHtmlSaveOptions : TestDataHelper
    {
        [Test]
        public static void SaveHtmlWithMetafileFormat()
        {
            //ExStart:SaveHtmlWithMetafileFormat
            Document doc = new Document(LoadingSavingDir + "Document.docx");
            HtmlSaveOptions options = new HtmlSaveOptions();
            options.MetafileFormat = HtmlMetafileFormat.EmfOrWmf;

            doc.Save(ArtifactsDir + "SaveHtmlWithMetafileFormat.html", options);
            //ExEnd:SaveHtmlWithMetafileFormat
        }

        [Test]
        public static void ImportExportSvgInHtml()
        {
            //ExStart:ImportExportSVGinHTML
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);
            
            builder.Write("Here is an SVG image: ");
            builder.InsertHtml(
                @"<svg height='210' width='500'>
                <polygon points='100,10 40,198 190,78 10,78 160,198' 
                    style='fill:lime;stroke:purple;stroke-width:5;fill-rule:evenodd;' />
            </svg> ");

            HtmlSaveOptions options = new HtmlSaveOptions();
            options.MetafileFormat = HtmlMetafileFormat.Svg;

            doc.Save(ArtifactsDir + "ImportExportSvgInHtml.html", options);
            //ExEnd:ImportExportSVGinHTML
        }

        [Test]
        public static void SetCssClassNamePrefix()
        {
            //ExStart:SetCssClassNamePrefix
            Document doc = new Document(LoadingSavingDir + "Document.docx");

            HtmlSaveOptions saveOptions = new HtmlSaveOptions();
            saveOptions.CssStyleSheetType = CssStyleSheetType.External;
            saveOptions.CssClassNamePrefix = "pfx_";

            doc.Save(ArtifactsDir + "SetCssClassNamePrefix.html", saveOptions);
            //ExEnd:SetCssClassNamePrefix
        }

        [Test]
        public static void SetExportCidUrlsForMhtmlResources()
        {
            //ExStart:SetExportCidUrlsForMhtmlResources
            Document doc = new Document(LoadingSavingDir + "CidUrls.docx");

            HtmlSaveOptions saveOptions = new HtmlSaveOptions(SaveFormat.Mhtml);
            saveOptions.PrettyFormat = true;
            saveOptions.ExportCidUrlsForMhtmlResources = true;
            saveOptions.SaveFormat = SaveFormat.Mhtml;

            doc.Save(ArtifactsDir + "SetExportCidUrlsForMhtmlResources.mhtml", saveOptions);
            //ExEnd:SetExportCidUrlsForMhtmlResources
        }

        [Test]
        public static void SetResolveFontNames()
        {
            //ExStart:SetResolveFontNames
            Document doc = new Document(LoadingSavingDir + "Test File (docx).docx");

            HtmlSaveOptions saveOptions = new HtmlSaveOptions(SaveFormat.Html);
            saveOptions.PrettyFormat = true;
            saveOptions.ResolveFontNames = true;

            doc.Save(ArtifactsDir + "SetResolveFontNames.html", saveOptions);
            //ExEnd:SetResolveFontNames
        }
    }
}