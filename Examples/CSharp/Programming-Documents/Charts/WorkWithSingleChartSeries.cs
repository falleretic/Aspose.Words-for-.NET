﻿using Aspose.Words.Drawing;
using Aspose.Words.Drawing.Charts;

namespace Aspose.Words.Examples.CSharp.Programming_Documents.Working_with_Charts
{
    class WorkWithSingleChartSeries : TestDataHelper
    {
        public static void Run()
        {
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            Shape shape = builder.InsertChart(ChartType.Line, 432, 252);
            Chart chart = shape.Chart;

            //ExStart:WorkWithSingleChartSeries
            // Get first series
            ChartSeries series0 = chart.Series[0];

            // Get second series
            ChartSeries series1 = chart.Series[1];

            // Change first series name
            series0.Name = "My Name1";

            // Change second series name
            series1.Name = "My Name2";

            // You can also specify whether the line connecting the points on the chart shall be smoothed using Catmull-Rom splines
            series0.Smooth = true;
            series1.Smooth = true;
            //ExEnd:WorkWithSingleChartSeries

            //ExStart:ChartDataPoint 
            // Specifies whether by default the parent element shall inverts its colors if the value is negative
            series0.InvertIfNegative = true;

            // Set default marker symbol and size
            series0.Marker.Symbol = MarkerSymbol.Circle;
            series0.Marker.Size = 15;

            series1.Marker.Symbol = MarkerSymbol.Star;
            series1.Marker.Size = 10;
            //ExEnd:ChartDataPoint 

            doc.Save(ArtifactsDir + "SingleChartSeries.docx");
        }
    }
}