using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;
using NUnit.Framework;

namespace SiteExamples.File_Formats_and_Conversions.Save_Options
{
    internal class WorkingWithHtmlSaveOptions : SiteExamplesBase
    {
        [Test, Description("Shows how to add roundtrip information as -aw-* for CSS elements.")]
        public void ExportRoundtripInformation()
        {
            //ExStart:ExportRoundtripInformation
            Document doc = new Document(MyDir + "Rendering.docx");

            HtmlSaveOptions saveOptions = new HtmlSaveOptions();
            saveOptions.ExportRoundtripInformation = true;
            
            doc.Save(ArtifactsDir + "HtmlSaveOptions.ExportRoundtripInformation.html", saveOptions);
            //ExEnd:ExportRoundtripInformation
        }

        [Test, Description("Shows how to export fonts in Base64 encoding.")]
        public void ExportFontsAsBase64()
        {
            //ExStart:ExportFontsAsBase64
            Document doc = new Document(MyDir + "Rendering.docx");
            
            HtmlSaveOptions saveOptions = new HtmlSaveOptions();
            saveOptions.ExportFontsAsBase64 = true;
            
            doc.Save(ArtifactsDir + "HtmlSaveOptions.ExportFontsAsBase64.html", saveOptions);
            //ExEnd:ExportFontsAsBase64
        }

        [Test, Description("Shows how to export CSS and fonts in external folder.")]
        public void ExportResources()
        {
            //ExStart:ExportResourcesUsingHtmlSaveOptions
            Document doc = new Document(MyDir + "Rendering.docx");
            
            HtmlSaveOptions saveOptions = new HtmlSaveOptions();
            saveOptions.CssStyleSheetType = CssStyleSheetType.External;
            saveOptions.ExportFontResources = true;
            saveOptions.ResourceFolder = ArtifactsDir + "Resources";
            saveOptions.ResourceFolderAlias = "http://example.com/resources";
            
            doc.Save(ArtifactsDir + "HtmlSaveOptions.ExportResourcesUsingHtmlSaveOptions.html", saveOptions);
            //ExEnd:ExportResourcesUsingHtmlSaveOptions
        }

        [Test, Description("Shows how to define metafile format when exporting to HTML.")]
        public void SaveHtmlWithMetafileFormat()
        {
            //ExStart:SaveHtmlWithMetafileFormat
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            builder.Write("Here is an image as is: ");
            builder.InsertHtml(
                @"<img src=""data:image/png;base64,
                    iVBORw0KGgoAAAANSUhEUgAAAAoAAAAKCAYAAACNMs+9AAAABGdBTUEAALGP
                    C/xhBQAAAAlwSFlzAAALEwAACxMBAJqcGAAAAAd0SU1FB9YGARc5KB0XV+IA
                    AAAddEVYdENvbW1lbnQAQ3JlYXRlZCB3aXRoIFRoZSBHSU1Q72QlbgAAAF1J
                    REFUGNO9zL0NglAAxPEfdLTs4BZM4DIO4C7OwQg2JoQ9LE1exdlYvBBeZ7jq
                    ch9//q1uH4TLzw4d6+ErXMMcXuHWxId3KOETnnXXV6MJpcq2MLaI97CER3N0
                    vr4MkhoXe0rZigAAAABJRU5ErkJggg=="" alt=""Red dot"" />");
            
            HtmlSaveOptions saveOptions = new HtmlSaveOptions();
            saveOptions.MetafileFormat = HtmlMetafileFormat.EmfOrWmf;

            doc.Save(ArtifactsDir + "HtmlSaveOptions.SaveHtmlWithMetafileFormat.html", saveOptions);
            //ExEnd:SaveHtmlWithMetafileFormat
        }

        [Test, Description("Shows how to define metafile format when exporting to HTML.")]
        public void ImportExportSvgInHtml()
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

            HtmlSaveOptions saveOptions = new HtmlSaveOptions();
            saveOptions.MetafileFormat = HtmlMetafileFormat.Svg;

            doc.Save(ArtifactsDir + "HtmlSaveOptions.ImportExportSvgInHtml.html", saveOptions);
            //ExEnd:ImportExportSVGinHTML
        }

        [Test, Description("Shows how to specify a prefix added to all CSS class names.")]
        public void CssClassNamePrefix()
        {
            //ExStart:CssClassNamePrefix
            Document doc = new Document(MyDir + "Rendering.docx");

            HtmlSaveOptions saveOptions = new HtmlSaveOptions();
            saveOptions.CssStyleSheetType = CssStyleSheetType.External;
            saveOptions.CssClassNamePrefix = "pfx_";

            doc.Save(ArtifactsDir + "HtmlSaveOptions.CssClassNamePrefix.html", saveOptions);
            //ExEnd:CssClassNamePrefix
        }

        [Test, Description("Shows how to set references to resource files as CID URLs.")]
        public void ExportCidUrlsForMhtmlResources()
        {
            //ExStart:ExportCidUrlsForMhtmlResources
            Document doc = new Document(MyDir + "Content-ID.docx");

            HtmlSaveOptions saveOptions = new HtmlSaveOptions(SaveFormat.Mhtml);
            saveOptions.PrettyFormat = true;
            saveOptions.ExportCidUrlsForMhtmlResources = true;
            saveOptions.SaveFormat = SaveFormat.Mhtml;

            doc.Save(ArtifactsDir + "HtmlSaveOptions.ExportCidUrlsForMhtmlResources.mhtml", saveOptions);
            //ExEnd:ExportCidUrlsForMhtmlResources
        }

        [Test, Description("Shows how to resolve fonts based on available font family.")]
        public void ResolveFontNames()
        {
            //ExStart:ResolveFontNames
            Document doc = new Document(MyDir + "Missing font.docx");

            HtmlSaveOptions saveOptions = new HtmlSaveOptions(SaveFormat.Html);
            saveOptions.PrettyFormat = true;
            saveOptions.ResolveFontNames = true;

            doc.Save(ArtifactsDir + "HtmlSaveOptionsEx.SetResolveFontNames.html", saveOptions);
            //ExEnd:ResolveFontNames
        }

        [Test, Description("Shows how to export form fields as plain text.")]
        public void ExportTextInputFormFieldAsText()
        {
            //ExStart:ExportTextInputFormFieldAsText
            Document doc = new Document(MyDir + "Rendering.docx");

            string imagesDir = Path.Combine(ArtifactsDir, "Images");

            // The folder specified needs to exist and should be empty
            if (Directory.Exists(imagesDir))
                Directory.Delete(imagesDir, true);

            Directory.CreateDirectory(imagesDir);

            // Set an option to export form fields as plain text, not as HTML input elements.
            HtmlSaveOptions saveOptions = new HtmlSaveOptions(SaveFormat.Html);
            saveOptions.ExportTextInputFormFieldAsText = true;
            saveOptions.ImagesFolder = imagesDir;

            doc.Save(ArtifactsDir + "HtmlSaveOptions.ExportTextInputFormFieldAsText.html", saveOptions);
            //ExEnd:ExportTextInputFormFieldAsText
        }
    }
}