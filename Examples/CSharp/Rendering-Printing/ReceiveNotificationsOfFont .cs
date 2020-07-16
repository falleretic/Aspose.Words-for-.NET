using System;
using Aspose.Words.Fonts;
using NUnit.Framework;

namespace Aspose.Words.Examples.CSharp
{
    class ReceiveNotificationsOfFont : TestDataHelper
    {
        [Test]
        public static void ReceiveNotificationsOfFonts()
        {
            //ExStart:ReceiveNotificationsOfFonts
            Document doc = new Document(RenderingPrintingDir + "Rendering.doc");

            FontSettings fontSettings = new FontSettings();

            // We can choose the default font to use in the case of any missing fonts.
            fontSettings.SubstitutionSettings.DefaultFontSubstitution.DefaultFontName = "Arial";
            // For testing we will set Aspose.Words to look for fonts only in a folder which doesn't exist. Since Aspose.Words won't
            // Find any fonts in the specified directory, then during rendering the fonts in the document will be subsuited with the default
            // Font specified under FontSettings.DefaultFontName. We can pick up on this subsuition using our callback
            fontSettings.SetFontsFolder(string.Empty, false);

            // Create a new class implementing IWarningCallback which collect any warnings produced during document save
            HandleDocumentWarnings callback = new HandleDocumentWarnings();

            doc.WarningCallback = callback;
            // Set font settings
            doc.FontSettings = fontSettings;
            
            doc.Save(ArtifactsDir + "ReceiveNotificationsOfFonts.pdf");
            //ExEnd:ReceiveNotificationsOfFonts
        }

        [Test]
        public static void ReceiveWarningNotification()
        {
            //ExStart:ReceiveWarningNotification
            Document doc = new Document(RenderingPrintingDir + "Rendering.doc");
            
            // When you call UpdatePageLayout the document is rendered in memory. Any warnings that occured during rendering
            // Are stored until the document save and then sent to the appropriate WarningCallback
            doc.UpdatePageLayout();

            // Create a new class implementing IWarningCallback and assign it to the PdfSaveOptions class
            HandleDocumentWarnings callback = new HandleDocumentWarnings();

            doc.WarningCallback = callback;
            
            // Even though the document was rendered previously, any save warnings are notified to the user during document save
            doc.Save(ArtifactsDir + "ReceiveWarningNotification.pdf");
            //ExEnd:ReceiveWarningNotification  
        }

        //ExStart:HandleDocumentWarnings
        public class HandleDocumentWarnings : IWarningCallback
        {
            /// <summary>
            /// Our callback only needs to implement the "Warning" method. This method is called whenever there is a
            /// Potential issue during document procssing. The callback can be set to listen for warnings generated during document
            /// Load and/or document save.
            /// </summary>
            public void Warning(WarningInfo info)
            {
                // We are only interested in fonts being substituted
                if (info.WarningType == WarningType.FontSubstitution)
                {
                    Console.WriteLine("Font substitution: " + info.Description);
                }
            }
        }
        //ExEnd:HandleDocumentWarnings
    }
}