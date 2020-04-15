﻿using Aspose.Words.Drawing;
using Aspose.Words.Drawing.Charts;
using NUnit.Framework;

namespace Aspose.Words.Examples.CSharp.Programming_Documents.Working_with_Charts
{
    class CreateChartUsingShape : TestDataHelper
    {
        [Test]
        public static void Run()
        {
            //ExStart:CreateChartUsingShape
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            Shape shape = builder.InsertChart(ChartType.Line, 432, 252);
            Chart chart = shape.Chart;

            // Determines whether the title shall be shown for this chart. Default is true
            chart.Title.Show = true;

            // Setting chart Title
            chart.Title.Text = "Sample Line Chart Title";

            // Determines whether other chart elements shall be allowed to overlap title
            chart.Title.Overlay = false;

            // Please note if null or empty value is specified as title text, auto generated title will be shown

            // Determines how legend shall be shown for this chart
            chart.Legend.Position = LegendPosition.Left;
            chart.Legend.Overlay = true;
            
            doc.Save(ArtifactsDir + "SimpleLineChart.docx");
            //ExEnd:CreateChartUsingShape
        }
    }
}