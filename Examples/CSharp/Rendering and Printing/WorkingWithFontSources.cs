using System;
using System.Collections;
using Aspose.Words.Fonts;
using NUnit.Framework;

namespace Aspose.Words.Examples.CSharp
{
    class WorkingWithFontSources : TestDataHelper
    {
        [Test]
        public static void GetListOfAvailableFonts()
        {
            //ExStart:GetListOfAvailableFonts
            FontSettings fontSettings = new FontSettings();
            ArrayList fontSources = new ArrayList(fontSettings.GetFontsSources());

            // Add a new folder source which will instruct Aspose.Words to search the following folder for fonts
            FolderFontSource folderFontSource = new FolderFontSource(MailMergeDir, true);
            // Add the custom folder which contains our fonts to the list of existing font sources
            fontSources.Add(folderFontSource);

            // Convert the ArrayList of source back into a primitive array of FontSource objects
            FontSourceBase[] updatedFontSources = (FontSourceBase[]) fontSources.ToArray(typeof(FontSourceBase));

            foreach (PhysicalFontInfo fontInfo in updatedFontSources[0].GetAvailableFonts())
            {
                Console.WriteLine("FontFamilyName : " + fontInfo.FontFamilyName);
                Console.WriteLine("FullFontName  : " + fontInfo.FullFontName);
                Console.WriteLine("Version  : " + fontInfo.Version);
                Console.WriteLine("FilePath : " + fontInfo.FilePath);
            }
            //ExEnd:GetListOfAvailableFonts
        }
    }
}