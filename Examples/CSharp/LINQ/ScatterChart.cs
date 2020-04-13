using Aspose.Words.Reporting;

namespace Aspose.Words.Examples.CSharp.LINQ
{
    class ScatterChart : TestDataHelper
    {
        public static void Run()
        {
            //ExStart:ScatterChart
            Document doc = new Document(LinqDir + "ScatterChart.docx");

            // Create a Reporting Engine
            ReportingEngine engine = new ReportingEngine();
            // Execute the build report
            engine.BuildReport(doc, Common.GetContracts(), "contracts");

            doc.Save(ArtifactsDir + "ScatterChart.docx");
            //ExEnd:ScatterChart
        }
    }
}