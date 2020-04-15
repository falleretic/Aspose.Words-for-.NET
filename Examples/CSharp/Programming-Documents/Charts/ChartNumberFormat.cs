using Aspose.Words.Drawing;
using Aspose.Words.Drawing.Charts;
using NUnit.Framework;

namespace Aspose.Words.Examples.CSharp.Programming_Documents.Working_with_Charts
{
    class ChartNumberFormat : TestDataHelper
    {
        [Test]
        public static void FormatNumberOfDataLabel()
        {
            //ExStart:FormatNumberofDataLabel
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Add chart with default data
            Shape shape = builder.InsertChart(ChartType.Line, 432, 252);
            Chart chart = shape.Chart;
            chart.Title.Text = "Data Labels With Different Number Format";

            // Delete default generated series
            chart.Series.Clear();

            // Add new series
            ChartSeries series0 = chart.Series.Add("AW Series 0", new string[] { "AW0", "AW1", "AW2" },
                new double[] { 2.5, 1.5, 3.5 });

            // Add DataLabel to the first point of the first series
            ChartDataLabel chartDataLabel0 = series0.DataLabels.Add(0);
            chartDataLabel0.ShowValue = true;

            // Set currency format code
            chartDataLabel0.NumberFormat.FormatCode = "\"$\"#,##0.00";

            ChartDataLabel chartDataLabel1 = series0.DataLabels.Add(1);
            chartDataLabel1.ShowValue = true;

            // Set date format code
            chartDataLabel1.NumberFormat.FormatCode = "d/mm/yyyy";

            ChartDataLabel chartDataLabel2 = series0.DataLabels.Add(2);
            chartDataLabel2.ShowValue = true;

            // Set percentage format code
            chartDataLabel2.NumberFormat.FormatCode = "0.00%";

            // Or you can set format code to be linked to a source cell,
            // in this case NumberFormat will be reset to general and inherited from a source cell
            chartDataLabel2.NumberFormat.IsLinkedToSource = true;

            doc.Save(ArtifactsDir + "NumberFormat_DataLabel.docx");
            //ExEnd:FormatNumberofDataLabel
        }
    }
}