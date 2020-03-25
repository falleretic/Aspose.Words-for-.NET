using System;
using Aspose.Words.Drawing;
using Aspose.Words.Drawing.Charts;

namespace Aspose.Words.Examples.CSharp.Programming_Documents.Working_with_Charts
{
    class WorkWithSingleChartDataPoint : TestDataHelper
    {
        public static void Run()
        {
            //ExStart:WorkWithSingleChartDataPoint
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            Shape shape = builder.InsertChart(ChartType.Line, 432, 252);
            Chart chart = shape.Chart;

            // Get first series
            ChartSeries series0 = chart.Series[0];
            // Get second series
            ChartSeries series1 = chart.Series[1];
            ChartDataPointCollection dataPointCollection = series0.DataPoints;

            // Add data point to the first and second point of the first series
            ChartDataPoint dataPoint00 = dataPointCollection.Add(0);
            ChartDataPoint dataPoint01 = dataPointCollection.Add(1);

            // Set explosion
            dataPoint00.Explosion = 50;

            // Set marker symbol and size
            dataPoint00.Marker.Symbol = MarkerSymbol.Circle;
            dataPoint00.Marker.Size = 15;

            dataPoint01.Marker.Symbol = MarkerSymbol.Diamond;
            dataPoint01.Marker.Size = 20;

            // Add data point to the third point of the second series
            ChartDataPoint dataPoint12 = series1.DataPoints.Add(2);
            dataPoint12.InvertIfNegative = true;
            dataPoint12.Marker.Symbol = MarkerSymbol.Star;
            dataPoint12.Marker.Size = 20;

            doc.Save(ArtifactsDir + "SingleChartDataPoint.docx");
            //ExEnd:WorkWithSingleChartDataPoint

            Console.WriteLine("\nSingle line chart created successfully.");
        }
    }
}