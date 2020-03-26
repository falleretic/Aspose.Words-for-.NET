using System;

namespace Aspose.Words.Examples.CSharp.Programming_Documents.Working_With_Document
{
    class SetupLanguagePreferences : TestDataHelper
    {
        public static void Run()
        {
            AddJapaneseAsEditingLanguages();
            SetRussianAsDefaultEditingLanguage();
        }

        private static void AddJapaneseAsEditingLanguages()
        {
            //ExStart:AddJapaneseAsEditinglanguages
            LoadOptions loadOptions = new LoadOptions();
            loadOptions.LanguagePreferences.AddEditingLanguage(EditingLanguage.Japanese);

            Document doc = new Document(DocumentDir + "languagepreferences.docx", loadOptions);

            int localeIdFarEast = doc.Styles.DefaultFont.LocaleIdFarEast;
            Console.WriteLine(
                localeIdFarEast == (int) EditingLanguage.Japanese
                    ? "The document either has no any FarEast language set in defaults or it was set to Japanese originally."
                    : "The document default FarEast language was set to another than Japanese language originally, so it is not overridden.");
            //ExEnd:AddJapaneseAsEditinglanguages
        }

        private static void SetRussianAsDefaultEditingLanguage()
        {
            //ExStart:SetRussianAsDefaultEditingLanguage
            LoadOptions loadOptions = new LoadOptions();
            loadOptions.LanguagePreferences.DefaultEditingLanguage = EditingLanguage.Russian;

            Document doc = new Document(DocumentDir + @"languagepreferences.docx", loadOptions);

            int localeId = doc.Styles.DefaultFont.LocaleId;
            Console.WriteLine(
                localeId == (int) EditingLanguage.Russian
                    ? "The document either has no any language set in defaults or it was set to Russian originally."
                    : "The document default language was set to another than Russian language originally, so it is not overridden.");
            //ExEnd:SetRussianAsDefaultEditingLanguage
        }
    }
}