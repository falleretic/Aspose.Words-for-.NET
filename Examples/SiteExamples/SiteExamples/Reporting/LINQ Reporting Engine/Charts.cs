using System.Collections.Generic;
using Aspose.Words;
using Aspose.Words.Reporting;
using NUnit.Framework;
using SiteExamples.Reporting.LINQ_Reporting_Engine.Helpers.Data_Source_Objects;

namespace SiteExamples.Reporting.LINQ_Reporting_Engine
{
    internal class Charts : SiteExamplesBase
    {
        [Test]
        public static void CreateBubbleChart()
        {
            //ExStart:BubbleChart
            Document doc = new Document(MyDir + "Reporting engine template - Bubble chart.docx");

            // Create a Reporting Engine
            ReportingEngine engine = new ReportingEngine();
            // Execute the build report
            engine.BuildReport(doc, Helpers.Common.GetContracts(), "contracts");
            
            doc.Save(ArtifactsDir + "ReportingEngine.CreateBubbleChart.docx");
            //ExEnd:BubbleChart
        }

        [Test]
        public static void SetChartSeriesNameDynamically()
        {
            //ExStart:SetChartSeriesNameDynamically
            List<PointData> data = new List<PointData>
            {
                new PointData { Time = "12:00:00 AM", Flow = 10, Rainfall = 2 },
                new PointData { Time = "01:00:00 AM", Flow = 15, Rainfall = 4 },
                new PointData { Time = "02:00:00 AM", Flow = 23, Rainfall = 7 }
            };

            List<string> seriesNames = new List<string> { "Flow", "Rainfall" };

            Document doc = new Document(MyDir + "Reporting engine template - Chart.docx");

            ReportingEngine engine = new ReportingEngine();
            engine.BuildReport(doc, new object[] { data, seriesNames }, new[] { "data", "seriesNames" });

            doc.Save(ArtifactsDir + "ReportingEngine.SetChartSeriesNameDynamically.docx");
            //ExEnd:SetChartSeriesNameDynamically
        }

        [Test]
        public static void ChartWithFilteringGroupingOrdering()
        {
            //ExStart:ChartWithFilteringGroupingOrdering
            Document doc = new Document(MyDir + "Reporting engine template - Chart with filtering.docx");

            ReportingEngine engine = new ReportingEngine();
            engine.BuildReport(doc, Helpers.Common.GetContracts(), "contracts");

            doc.Save(ArtifactsDir + "ReportingEngine.ChartWithFilteringGroupingOrdering.docx");
            //ExEnd:ChartWithFilteringGroupingOrdering
        }

        [Test]
        public static void PieChart()
        {
            //ExStart:PieChart
            Document doc = new Document(MyDir + "Reporting engine template - Pie chart.docx");

            ReportingEngine engine = new ReportingEngine();
            engine.BuildReport(doc, Helpers.Common.GetManagers(), "managers");

            doc.Save(ArtifactsDir + "ReportingEngine.PieChart.docx");
            //ExEnd:PieChart
        }

        [Test]
        public static void ScatterChart()
        {
            //ExStart:ScatterChart
            Document doc = new Document(MyDir + "Reporting engine template - Scatter chart.docx");

            ReportingEngine engine = new ReportingEngine();
            engine.BuildReport(doc, Helpers.Common.GetContracts(), "contracts");

            doc.Save(ArtifactsDir + "ReportingEngine.ScatterChart.docx");
            //ExEnd:ScatterChart
        }
    }
}