using System.IO;
using Aspose.Words.Saving;
using NUnit.Framework;

namespace Aspose.Words.Examples.CSharp
{
    class HtmlSaveOptionsEx : TestDataHelper
    {
        [Test]
        public static void ConvertDocumentToEpub()
        {
            //ExStart:ConvertDocumentToEPUB
            Document doc = new Document(LoadingSavingDir + "Rendering.docx");

            // Create a new instance of HtmlSaveOptions. This object allows us to set options that control
            // how the output document is saved
            HtmlSaveOptions saveOptions = new HtmlSaveOptions();
            // Specify the desired encoding
            saveOptions.Encoding = System.Text.Encoding.UTF8;
            // Specify at what elements to split the internal HTML at. This creates a new HTML within the EPUB 
            // which allows you to limit the size of each HTML part. This is useful for readers which cannot read 
            // HTML files greater than a certain size e.g 300kb.
            saveOptions.DocumentSplitCriteria = DocumentSplitCriteria.HeadingParagraph;
            // Specify that we want to export document properties
            saveOptions.ExportDocumentProperties = true;
            // Specify that we want to save in EPUB format
            saveOptions.SaveFormat = SaveFormat.Epub;

            // Export the document as an EPUB file
            doc.Save(ArtifactsDir + "HtmlSaveOptionsEx.ConvertDocumentToEPUB.epub", saveOptions);
            //ExEnd:ConvertDocumentToEPUB
        }

        [Test]
        public static void ExportRoundtripInformation()
        {
            //ExStart:ConvertDocumentToHtmlWithRoundtrip
            // Load the document from disk.
            Document doc = new Document(LoadingSavingDir + "Rendering.docx");

            HtmlSaveOptions options = new HtmlSaveOptions();
            // HtmlSaveOptions.ExportRoundtripInformation property specifies
            // Whether to write the roundtrip information when saving to HTML, MHTML or EPUB.
            // Default value is true for HTML and false for MHTML and EPUB.
            options.ExportRoundtripInformation = true;
            
            doc.Save(ArtifactsDir + "HtmlSaveOptionsEx.ExportRoundtripInformation.html", options);
            //ExEnd:ConvertDocumentToHtmlWithRoundtrip
        }

        [Test]
        public static void SplitDocumentByHeadingsHtml()
        {
            //ExStart:SplitDocumentByHeadingsHtml
            // Open a Word document
            Document doc = new Document(LoadingSavingDir + "Rendering.docx");
 
            HtmlSaveOptions options = new HtmlSaveOptions();
            // Split a document into smaller parts, in this instance split by heading
            options.DocumentSplitCriteria = DocumentSplitCriteria.HeadingParagraph;
 
            // Save the output file
            doc.Save(ArtifactsDir + "HtmlSaveOptionsEx.SplitDocumentByHeadings.html", options);
            //ExEnd:SplitDocumentByHeadingsHtml
        }

        [Test]
        public static void SplitDocumentBySectionsHtml()
        {
            // Open a Word document
            Document doc = new Document(LoadingSavingDir + "Rendering.docx");
 
            //ExStart:SplitDocumentBySectionsHtml
            HtmlSaveOptions options = new HtmlSaveOptions();
            options.DocumentSplitCriteria = DocumentSplitCriteria.SectionBreak;
            //ExEnd:SplitDocumentBySectionsHtml
            
            // Save the output file
            doc.Save(ArtifactsDir + "HtmlSaveOptionsEx.SplitDocumentBySections.html", options);
        }

        [Test]
        public static void ExportFontsAsBase64()
        {
            //ExStart:ExportFontsAsBase64
            Document doc = new Document(LoadingSavingDir + "Rendering.docx");
            
            HtmlSaveOptions saveOptions = new HtmlSaveOptions();
            saveOptions.ExportFontResources = true;
            saveOptions.ExportFontsAsBase64 = true;
            
            doc.Save(ArtifactsDir + "HtmlSaveOptionsEx.ExportFontsAsBase64.html", saveOptions);
            //ExEnd:ExportFontsAsBase64
        }

        [Test]
        public static void ExportResources()
        {
            //ExStart:ExportResourcesUsingHtmlSaveOptions
            Document doc = new Document(LoadingSavingDir + "Rendering.docx");
            
            HtmlSaveOptions saveOptions = new HtmlSaveOptions();
            saveOptions.CssStyleSheetType = CssStyleSheetType.External;
            saveOptions.ExportFontResources = true;
            saveOptions.ResourceFolder = ArtifactsDir + "Resources";
            saveOptions.ResourceFolderAlias = "http://example.com/resources";
            
            doc.Save(ArtifactsDir + "HtmlSaveOptionsEx.ExportResourcesUsingHtmlSaveOptions.html", saveOptions);
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

            doc.Save(ArtifactsDir + "HtmlSaveOptionsEx.SaveHtmlWithMetafileFormat.html", options);
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

            doc.Save(ArtifactsDir + "HtmlSaveOptionsEx.ImportExportSvgInHtml.html", options);
            //ExEnd:ImportExportSVGinHTML
        }

        [Test]
        public static void SetCssClassNamePrefix()
        {
            //ExStart:SetCssClassNamePrefix
            Document doc = new Document(LoadingSavingDir + "Rendering.docx");

            HtmlSaveOptions saveOptions = new HtmlSaveOptions();
            saveOptions.CssStyleSheetType = CssStyleSheetType.External;
            saveOptions.CssClassNamePrefix = "pfx_";

            doc.Save(ArtifactsDir + "HtmlSaveOptionsEx.SetCssClassNamePrefix.html", saveOptions);
            //ExEnd:SetCssClassNamePrefix
        }

        [Test]
        public static void SetExportCidUrlsForMhtmlResources()
        {
            //ExStart:SetExportCidUrlsForMhtmlResources
            Document doc = new Document(LoadingSavingDir + "Content-ID.docx");

            HtmlSaveOptions saveOptions = new HtmlSaveOptions(SaveFormat.Mhtml);
            saveOptions.PrettyFormat = true;
            saveOptions.ExportCidUrlsForMhtmlResources = true;
            saveOptions.SaveFormat = SaveFormat.Mhtml;

            doc.Save(ArtifactsDir + "HtmlSaveOptionsEx.SetExportCidUrlsForMhtmlResources.mhtml", saveOptions);
            //ExEnd:SetExportCidUrlsForMhtmlResources
        }

        [Test]
        public static void SetResolveFontNames()
        {
            //ExStart:SetResolveFontNames
            Document doc = new Document(LoadingSavingDir + "Missing font.docx");

            HtmlSaveOptions saveOptions = new HtmlSaveOptions(SaveFormat.Html);
            saveOptions.PrettyFormat = true;
            saveOptions.ResolveFontNames = true;

            doc.Save(ArtifactsDir + "HtmlSaveOptionsEx.SetResolveFontNames.html", saveOptions);
            //ExEnd:SetResolveFontNames
        }

        [Test]
        public static void SpecifySaveOption()
        {
            //ExStart:SpecifySaveOption
            Document doc = new Document(LoadingSavingDir + "Rendering.docx");

            // This is the directory we want the exported images to be saved to
            string imagesDir = Path.Combine(ArtifactsDir, "Images");

            // The folder specified needs to exist and should be empty
            if (Directory.Exists(imagesDir))
                Directory.Delete(imagesDir, true);

            Directory.CreateDirectory(imagesDir);

            // Set an option to export form fields as plain text, not as HTML input elements
            HtmlSaveOptions options = new HtmlSaveOptions(SaveFormat.Html);
            options.ExportTextInputFormFieldAsText = true;
            options.ImagesFolder = imagesDir;

            doc.Save(ArtifactsDir + "HtmlSaveOptionsEx.SpecifySaveOption.html", options);
            //ExEnd:SpecifySaveOption
        }
    }
}
