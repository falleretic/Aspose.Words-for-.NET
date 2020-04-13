using Aspose.Words.Reporting;

namespace Aspose.Words.Examples.CSharp.LINQ
{
    class BuildOptions : TestDataHelper
    {
        public static void Run()
        {
            RemoveEmptyParagraphs();
        }

        public static void RemoveEmptyParagraphs()
        {
            //ExStart:RemoveEmptyParagraphs
            Document doc = new Document(LinqDir + "template_cleanup.docx");

            // Create a Reporting Engine
            ReportingEngine engine = new ReportingEngine();
            engine.Options = ReportBuildOptions.RemoveEmptyParagraphs;
            engine.BuildReport(doc, Common.GetManagers(), "managers");

            doc.Save(ArtifactsDir + "RemoveEmptyParagraphs.docx");
            //ExEnd:RemoveEmptyParagraphs
        }
    }
}