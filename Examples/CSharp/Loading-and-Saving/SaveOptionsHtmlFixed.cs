﻿using Aspose.Words.Saving;

namespace Aspose.Words.Examples.CSharp.Loading_Saving
{
    class SaveOptionsHtmlFixed : TestDataHelper
    {
        public static void Run()
        {
            UseFontFromTargetMachine();
            WriteAllCssRulesInSingleFile();
        }

        static void UseFontFromTargetMachine()
        {
            //ExStart:UseFontFromTargetMachine
            Document doc = new Document(LoadingSavingDir + "Test File (doc).doc");

            HtmlFixedSaveOptions options = new HtmlFixedSaveOptions();
            options.UseTargetMachineFonts = true;

            doc.Save(ArtifactsDir + "UseFontFromTargetMachine.html", options);
            //ExEnd:UseFontFromTargetMachine
        }

        static void WriteAllCssRulesInSingleFile()
        {
            //ExStart:WriteAllCSSrulesinSingleFile
            Document doc = new Document(LoadingSavingDir + "Test File (doc).doc");

            HtmlFixedSaveOptions options = new HtmlFixedSaveOptions();
            // Setting this property to true restores the old behavior (separate files) for compatibility with legacy code
            // Default value is false
            // All CSS rules are written into single file "styles.css
            options.SaveFontFaceCssSeparately = false;

            doc.Save(ArtifactsDir + "WriteAllCssRulesInSingleFile.html", options);
            //ExEnd:WriteAllCSSrulesinSingleFile
        }
    }
}