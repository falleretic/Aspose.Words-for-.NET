using Aspose.Words.Fonts;
using System.IO;
using System.Reflection;
using NUnit.Framework;

namespace Aspose.Words.Examples.CSharp.Rendering_and_Printing
{
    // ExStart:ResourceSteamFontSourceExample
    class ResourceSteamFontSourceExamples : TestDataHelper
    {
        [Test]
        public static void Run()
        {
            Document doc = new Document(RenderingPrintingDir + "Rendering.doc");
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