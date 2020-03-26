﻿using Aspose.Words.Drawing;
using Aspose.Words.Drawing.Charts;

namespace Aspose.Words.Examples.CSharp.Programming_Documents.Working_with_Charts
{
    class InsertScatterChart : TestDataHelper
    {
        public static void Run()
        {
            //ExStart:InsertScatterChart
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Insert Scatter chart
            Shape shape = builder.InsertChart(ChartType.Scatter, 432, 252);
            Chart chart = shape.Chart;

            // Use this overload to add series to any type of Scatter charts
            chart.Series.Add("AW Series 1", new double[] { 0.7, 1.8, 2.6 }, new double[] { 2.7, 3.2, 0.8 });

            doc.Save(ArtifactsDir + "TestInsertScatterChart.docx");
            //ExEnd:InsertScatterChart
        }
    }
}