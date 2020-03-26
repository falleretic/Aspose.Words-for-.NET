using Aspose.Words.Drawing;
using Aspose.Words.Drawing.Charts;

namespace Aspose.Words.Examples.CSharp.Programming_Documents.Working_with_Charts
{
    class WorkWithChartDataLabels : TestDataHelper
    {
        public static void Run()
        {
            WorkWithChartDataLabel();
            DefaultOptionsForDataLabels();
        }

        public static void WorkWithChartDataLabel()
        {
            //ExStart:WorkWithChartDataLabel
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            Shape shape = builder.InsertChart(ChartType.Bar, 432, 252);
            Chart chart = shape.Chart;

            // Get first series
            ChartSeries series0 = chart.Series[0];
            ChartDataLabelCollection dataLabelCollection = series0.DataLabels;

            // Add data label to the first and second point of the first series
            ChartDataLabel chartDataLabel00 = dataLabelCollection.Add(0);
            ChartDataLabel chartDataLabel01 = dataLabelCollection.Add(1);

            // Set properties
            chartDataLabel00.ShowLegendKey = true;

            // By default, when you add data labels to the data points in a pie chart, leader lines are displayed for data labels that are
            // positioned far outside the end of data points. Leader lines create a visual connection between a data label and its 
            // corresponding data point
            chartDataLabel00.ShowLeaderLines = true;

            chartDataLabel00.ShowCategoryName = false;
            chartDataLabel00.ShowPercentage = false;
            chartDataLabel00.ShowSeriesName = true;
            chartDataLabel00.ShowValue = true;
            chartDataLabel00.Separator = "/";
            chartDataLabel01.ShowValue = true;
            
            doc.Save(ArtifactsDir + "SimpleBarChart.docx");
            //ExEnd:WorkWithChartDataLabel
        }

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