using Aspose.Words.Reporting;
using NUnit.Framework;

namespace Aspose.Words.Examples.CSharp.LINQ
{
    class BubbleChart : TestDataHelper
    {
        public static void Run()
        {
            CreateBubbleChart();
        }

        [Test]
        public static void CreateBubbleChart()
        {
            //ExStart:BubbleChart
            Document doc = new Document(LinqDir + "BubbleChart.docx");

            // Create a Reporting Engine
            ReportingEngine engine = new ReportingEngine();
            // Execute the build report
            engine.BuildReport(doc, Common.GetContracts(), "contracts");
            
            doc.Save(ArtifactsDir + "CreateBubbleChart.docx");
            //ExEnd:BubbleChart
        }
    }
}