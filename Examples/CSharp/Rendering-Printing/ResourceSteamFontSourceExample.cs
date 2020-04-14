using Aspose.Words.Fonts;
using System.IO;
using System.Reflection;

namespace Aspose.Words.Examples.CSharp.Rendering_and_Printing
{
    // ExStart:ResourceSteamFontSourceExample
    class ResourceSteamFontSourceExamples : TestDataHelper
    {
        public static void Run()
        {
            Document doc = new Document(MailMergeDir + "Rendering.doc");
            // FontSettings.SetFontSources instead
            FontSettings.DefaultInstance.SetFontsSources(new FontSourceBase[]
                { new SystemFontSource(), new ResourceSteamFontSource() });

            doc.Save(ArtifactsDir + "Rendering.SetFontsFolders.pdf");
        }
    }

    internal class ResourceSteamFontSource : StreamFontSource
    {
        public override Stream OpenFontDataStream()
        {
            return Assembly.GetExecutingAssembly().GetManifestResourceStream("resourceName");
        }
    }
    // ExEnd:ResourceSteamFontSourceExample
}