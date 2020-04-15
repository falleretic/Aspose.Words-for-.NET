using Aspose.Words.Settings;
using NUnit.Framework;

namespace Aspose.Words.Examples.CSharp.Programming_Documents.Working_With_Document
{
    class SetViewOption : TestDataHelper
    {
        [Test]
        public static void Run()
        {
            //ExStart:SetViewOption
            Document doc = new Document(DocumentDir + "TestFile.doc");
            // Set view option
            doc.ViewOptions.ViewType = ViewType.PageLayout;
            doc.ViewOptions.ZoomPercent = 50;

            doc.Save(ArtifactsDir + "TestFile.SetZoom_out.doc");
            //ExEnd:SetViewOption
        }
    }
}