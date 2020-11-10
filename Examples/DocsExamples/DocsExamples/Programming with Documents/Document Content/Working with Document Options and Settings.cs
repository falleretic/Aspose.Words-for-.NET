using System;
using Aspose.Words;
using Aspose.Words.Settings;
using NUnit.Framework;

namespace DocsExamples.Programming_with_Documents.Document_Content
{
    class SpecificDocumentOptions : DocsExamplesBase
    {
        [Test]
        public static void OptimizeFor()
        {
            //ExStart:OptimizeFor
            Document doc = new Document(MyDir + "Document.docx");
            doc.CompatibilityOptions.OptimizeFor(MsWordVersion.Word2016);

            doc.Save(ArtifactsDir + "TestFile.docx");
            //ExEnd:OptimizeFor
        }

        [Test]
        public static void ShowGrammaticalAndSpellingErrors()
        {
            // ExStart: ShowGrammaticalAndSpellingErrors
            Document doc = new Document(MyDir + "Document.docx");

            doc.ShowGrammaticalErrors = true;
            doc.ShowSpellingErrors = true;

            doc.Save(ArtifactsDir + "Document.ShowErrorsInDocument.docx");
            // ExEnd: ShowGrammaticalAndSpellingErrors
        }

        [Test]
        public static void CleanupUnusedStylesandLists()
        {
            // ExStart:CleanupUnusedStylesandLists
            Document doc = new Document(MyDir + "Document.docx");

            CleanupOptions cleanupOptions = new CleanupOptions();
            cleanupOptions.UnusedLists = false;
            cleanupOptions.UnusedStyles = true;

            // Cleans unused styles and lists from the document depending on given CleanupOptions. 
            doc.Cleanup(cleanupOptions);

            doc.Save(ArtifactsDir + "Document.CleanupUnusedStylesandLists.docx");
            // ExEnd:CleanupUnusedStylesandLists
        }

        [Test]
        public static void CleanupDuplicateStyle()
        {
            // ExStart:CleanupDuplicateStyle
            Document doc = new Document(MyDir + "Document.docx");

            CleanupOptions options = new CleanupOptions();
            options.DuplicateStyle = true;

            // Cleans duplicate styles from the document. 
            doc.Cleanup(options);

            doc.Save(ArtifactsDir + "Document.CleanupDuplicateStyle_out.docx");
            // ExEnd:CleanupDuplicateStyle
        }

        [Test]
        public static void SetViewOption()
        {
            //ExStart:SetViewOption
            Document doc = new Document(MyDir + "Document.docx");
            // Set view option
            doc.ViewOptions.ViewType = ViewType.PageLayout;
            doc.ViewOptions.ZoomPercent = 50;

            doc.Save(ArtifactsDir + "TestFile.SetZoom_out.doc");
            //ExEnd:SetViewOption
        }

        [Test]
        public static void DocumentPageSetup()
        {
            //ExStart:DocumentPageSetup
            Document doc = new Document(MyDir + "Document.docx");

            // Set the layout mode for a section allowing to define the document grid behavior
            // Note that the Document Grid tab becomes visible in the Page Setup dialog of MS Word
            // if any Asian language is defined as editing language
            doc.FirstSection.PageSetup.LayoutMode = SectionLayoutMode.Grid;
            // Set the number of characters per line in the document grid
            doc.FirstSection.PageSetup.CharactersPerLine = 30;
            // Set the number of lines per page in the document grid
            doc.FirstSection.PageSetup.LinesPerPage = 10;

            doc.Save(ArtifactsDir + "Document.PageSetup.doc");
            //ExEnd:DocumentPageSetup
        }

        [Test]
        public static void AddJapaneseAsEditingLanguages()
        {
            //ExStart:AddJapaneseAsEditinglanguages
            LoadOptions loadOptions = new LoadOptions();
            // Set language preferences that will be used when document is loading.
            loadOptions.LanguagePreferences.AddEditingLanguage(EditingLanguage.Japanese);

            Document doc = new Document(MyDir + "No default editing language.docx", loadOptions);

            int localeIdFarEast = doc.Styles.DefaultFont.LocaleIdFarEast;
            Console.WriteLine(
                localeIdFarEast == (int) EditingLanguage.Japanese
                    ? "The document either has no any FarEast language set in defaults or it was set to Japanese originally."
                    : "The document default FarEast language was set to another than Japanese language originally, so it is not overridden.");
            //ExEnd:AddJapaneseAsEditinglanguages
        }

        [Test]
        public static void SetRussianAsDefaultEditingLanguage()
        {
            //ExStart:SetRussianAsDefaultEditingLanguage
            Aspose.Words.LoadOptions loadOptions = new Aspose.Words.LoadOptions();
            loadOptions.LanguagePreferences.DefaultEditingLanguage = EditingLanguage.Russian;

            Document doc = new Document(MyDir + "No default editing language.docx", loadOptions);

            int localeId = doc.Styles.DefaultFont.LocaleId;
            Console.WriteLine(
                localeId == (int) EditingLanguage.Russian
                    ? "The document either has no any language set in defaults or it was set to Russian originally."
                    : "The document default language was set to another than Russian language originally, so it is not overridden.");
            //ExEnd:SetRussianAsDefaultEditingLanguage
        }

        [Test]
        public static void SetPageSetupAndSectionFormatting()
        {
            //ExStart:DocumentBuilderSetPageSetupAndSectionFormatting
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Set page properties
            builder.PageSetup.Orientation = Orientation.Landscape;
            builder.PageSetup.LeftMargin = 50;
            builder.PageSetup.PaperSize = PaperSize.Paper10x14;

            doc.Save(ArtifactsDir + "DocumentBuilderSetPageSetupAndSectionFormatting.doc");
            //ExEnd:DocumentBuilderSetPageSetupAndSectionFormatting
        }
    }
}