using Aspose.Words.Drawing;
using Aspose.Words.Drawing.Charts;
using NUnit.Framework;

namespace Aspose.Words.Examples.CSharp.Programming_Documents.Working_with_Charts
{
    class WorkWithChartDataLabels : TestDataHelper
    {
        [Test]
        public static void WorkWithChartDataLabel()
        {
            //ExStart:WorkWithChartDataLabel
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            Shape shape = builder.InsertChart(ChartType.Bar, 432, 252);
            Chart chart = shape.Chart;

            // Get first series.
            ChartSeries series0 = shape.Chart.Series[0];

            ChartDataLabelCollection labels = series0.DataLabels;

            // Set properties.
            labels.ShowLegendKey = true;

            // By default, when you add data labels to the data points in a pie chart, leader lines are displayed for data labels that are
            // Positioned far outside the end of data points. Leader lines create a visual connection between a data label and its 
            // Corresponding data point.
            labels.ShowLeaderLines = true;

            labels.ShowCategoryName = false;
            labels.ShowPercentage = false;
            labels.ShowSeriesName = true;
            labels.ShowValue = true;
            labels.Separator = "/";
            labels.ShowValue = true;
            
            doc.Save(ArtifactsDir + "SimpleBarChart.docx");
            //ExEnd:WorkWithChartDataLabel
        }

        [Test]
        public static void DefaultOptionsForDataLabels()
        {
            //ExStart:DefaultOptionsForDataLabels
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            Shape shape = builder.InsertChart(ChartType.Pie, 432, 252);
            Chart chart = shape.Chart;
            chart.Series.Clear();

            ChartSeries series = chart.Series.Add("Series 1",
                new string[] { "Category1", "Category2", "Category3" },
                new double[] { 2.7, 3.2, 0.8 });

            ChartDataLabelCollection labels = series.DataLabels;
            labels.ShowPercentage = true;
            labels.ShowValue = true;
            labels.ShowLeaderLines = false;
            labels.Separator = " - ";

            doc.Save(ArtifactsDir + "Demo.docx");
            //ExEnd:DefaultOptionsForDataLabels
        }
    }
}