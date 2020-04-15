using Aspose.Words.Fonts;
using NUnit.Framework;

namespace Aspose.Words.Examples.CSharp.Rendering_and_Printing
{
    class SpecifyDefaultFontWhenRendering : TestDataHelper
    {
        [Test]
        public static void Run()
        {
            //ExStart:SpecifyDefaultFontWhenRendering
            Document doc = new Document(MailMergeDir + "Rendering.doc");

            FontSettings fontSettings = new FontSettings();
            // If the default font defined here cannot be found during rendering then
            // the closest font on the machine is used instead
            fontSettings.SubstitutionSettings.DefaultFontSubstitution.DefaultFontName = "Arial Unicode MS";
            // Set font settings
            doc.FontSettings = fontSettings;
            
            // Now the set default font is used in place of any missing fonts during any rendering calls
            doc.Save(ArtifactsDir + "SpecifyDefaultFontWhenRendering.pdf");
            //ExEnd:SpecifyDefaultFontWhenRendering
        }
    }
}