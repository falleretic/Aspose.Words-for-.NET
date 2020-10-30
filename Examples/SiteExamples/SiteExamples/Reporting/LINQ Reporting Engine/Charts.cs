﻿using System.Collections.Generic;
using Aspose.Words;
using Aspose.Words.Reporting;
using NUnit.Framework;
using SiteExamples.Reporting.LINQ_Reporting_Engine.Helpers;
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
            engine.BuildReport(doc, Helpers.Common.GetManagers(), "managers");
            
            doc.Save(ArtifactsDir + "ReportingEngine.CreateBubbleChart.docx");
            //ExEnd:BubbleChart
        }

        [Test]
        public static void SetChartSeriesNameDynamically()
        {
            //ExStart:SetChartSeriesNameDynamically
            Document doc = new Document(MyDir + "Reporting engine template - Chart.docx");

            ReportingEngine engine = new ReportingEngine();
            engine.BuildReport(doc, Common.GetManagers(), "managers");

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