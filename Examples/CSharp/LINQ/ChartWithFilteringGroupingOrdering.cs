using Aspose.Words.Reporting;

namespace Aspose.Words.Examples.CSharp.LINQ
{
    class ChartWithFilteringGroupingOrdering : TestDataHelper
    {
        public static void Run()
        {
            //ExStart:ChartWithFilteringGroupingOrdering
            Document doc = new Document(LinqDir + "ChartWithFilteringGroupingOrdering.docx");

            // Create a Reporting Engine
            ReportingEngine engine = new ReportingEngine();
            // Execute the build report
            engine.BuildReport(doc, Common.GetContracts(), "contracts");

            doc.Save(ArtifactsDir + "ChartWithFilteringGroupingOrdering.docx");
            //ExEnd:ChartWithFilteringGroupingOrdering
        }
    }
}