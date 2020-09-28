using System.IO;
using Aspose.Words.Saving;
using NUnit.Framework;

namespace Aspose.Words.Examples.CSharp.File_Formats_and_Conversions.Save_Options
{
    class WorkingWithHtmlSaveOptions : TestDataHelper
    {
        [Test]
        public static void ConvertDocumentToEpub()
        {
            //ExStart:ConvertDocumentToEpub
            Document doc = new Document(MyDir + "Rendering.docx");

            HtmlSaveOptions saveOptions = new HtmlSaveOptions();
            saveOptions.Encoding = System.Text.Encoding.UTF8;
            // Specify at what elements to split the internal HTML at. This creates a new HTML within the EPUB 
            // which allows you to limit the size of each HTML part. This is useful for readers which cannot read 
            // HTML files greater than a certain size e.g 300kb.
            saveOptions.DocumentSplitCriteria = DocumentSplitCriteria.HeadingParagraph;
            saveOptions.ExportDocumentProperties = true;
            saveOptions.SaveFormat = SaveFormat.Epub;

            doc.Save(ArtifactsDir + "HtmlSaveOptions.ConvertDocumentToEpub.epub", saveOptions);
            //ExEnd:ConvertDocumentToEpub
        }

        [Test]
        public static void ExportRoundtripInformation()
        {
            //ExStart:ExportRoundtripInformation
            Document doc = new Document(MyDir + "Rendering.docx");

            HtmlSaveOptions options = new HtmlSaveOptions();
            options.ExportRoundtripInformation = true;
            
            doc.Save(ArtifactsDir + "HtmlSaveOptions.ExportRoundtripInformation.html", options);
            //ExEnd:ExportRoundtripInformation
        }

        [Test]
        public static void ExportFontsAsBase64()
        {
            //ExStart:ExportFontsAsBase64
            Document doc = new Document(MyDir + "Rendering.docx");
            
            HtmlSaveOptions saveOptions = new HtmlSaveOptions();
            saveOptions.ExportFontResources = true;
            saveOptions.ExportFontsAsBase64 = true;
            
            doc.Save(ArtifactsDir + "HtmlSaveOptions.ExportFontsAsBase64.html", saveOptions);
            //ExEnd:ExportFontsAsBase64
        }

        [Test]
        public static void ExportResources()
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

        [Test]
        public static void SaveHtmlWithMetafileFormat()
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
            
            HtmlSaveOptions options = new HtmlSaveOptions();
            options.MetafileFormat = HtmlMetafileFormat.EmfOrWmf;

            doc.Save(ArtifactsDir + "HtmlSaveOptions.SaveHtmlWithMetafileFormat.html", options);
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

            doc.Save(ArtifactsDir + "HtmlSaveOptions.ImportExportSvgInHtml.html", options);
            //ExEnd:ImportExportSVGinHTML
        }

        [Test]
        public static void CssClassNamePrefix()
        {
            //ExStart:CssClassNamePrefix
            Document doc = new Document(MyDir + "Rendering.docx");

            HtmlSaveOptions saveOptions = new HtmlSaveOptions();
            saveOptions.CssStyleSheetType = CssStyleSheetType.External;
            saveOptions.CssClassNamePrefix = "pfx_";

            doc.Save(ArtifactsDir + "HtmlSaveOptions.SetCssClassNamePrefix.html", saveOptions);
            //ExEnd:CssClassNamePrefix
        }

        [Test]
        public static void ExportCidUrlsForMhtmlResources()
        {
            //ExStart:ExportCidUrlsForMhtmlResources
            Document doc = new Document(MyDir + "Content-ID.docx");

            HtmlSaveOptions saveOptions = new HtmlSaveOptions(SaveFormat.Mhtml);
            saveOptions.PrettyFormat = true;
            saveOptions.ExportCidUrlsForMhtmlResources = true;
            saveOptions.SaveFormat = SaveFormat.Mhtml;

            doc.Save(ArtifactsDir + "HtmlSaveOptions.SetExportCidUrlsForMhtmlResources.mhtml", saveOptions);
            //ExEnd:ExportCidUrlsForMhtmlResources
        }

        [Test]
        public static void ResolveFontNames()
        {
            //ExStart:ResolveFontNames
            Document doc = new Document(MyDir + "Missing font.docx");

            HtmlSaveOptions saveOptions = new HtmlSaveOptions(SaveFormat.Html);
            saveOptions.PrettyFormat = true;
            saveOptions.ResolveFontNames = true;

            doc.Save(ArtifactsDir + "HtmlSaveOptionsEx.SetResolveFontNames.html", saveOptions);
            //ExEnd:ResolveFontNames
        }

        [Test]
        public static void SpecifySaveOption()
        {
            //ExStart:SpecifySaveOption
            Document doc = new Document(MyDir + "Rendering.docx");

            string imagesDir = Path.Combine(ArtifactsDir, "Images");

            // The folder specified needs to exist and should be empty
            if (Directory.Exists(imagesDir))
                Directory.Delete(imagesDir, true);

            Directory.CreateDirectory(imagesDir);

            // Set an option to export form fields as plain text, not as HTML input elements.
            HtmlSaveOptions options = new HtmlSaveOptions(SaveFormat.Html);
            options.ExportTextInputFormFieldAsText = true;
            options.ImagesFolder = imagesDir;

            doc.Save(ArtifactsDir + "HtmlSaveOptions.SpecifySaveOption.html", options);
            //ExEnd:SpecifySaveOption
        }
    }
}
