using System;
using Aspose.Words.Drawing;
using Aspose.Words.Drawing.Charts;

namespace Aspose.Words.Examples.CSharp.Programming_Documents.Working_with_Charts
{
    class InsertAreaChart : TestDataHelper
    {
        public static void Run()
        {
            //ExStart:InsertAreaChart
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Insert Area chart
            Shape shape = builder.InsertChart(ChartType.Area, 432, 252);
            Chart chart = shape.Chart;

            // Use this overload to add series to any type of Area, Radar and Stock charts
            chart.Series.Add("AW Series 1", new []
            {
                new DateTime(2002, 05, 01),
                new DateTime(2002, 06, 01),
                new DateTime(2002, 07, 01),
                new DateTime(2002, 08, 01),
                new DateTime(2002, 09, 01)
            }, 
                new double[] { 32, 32, 28, 12, 15 });
            
            doc.Save(ArtifactsDir + "TestInsertAreaChart.docx");
            //ExEnd:InsertAreaChart

            Console.WriteLine("\nScatter chart created successfully.");
        }
    }
}