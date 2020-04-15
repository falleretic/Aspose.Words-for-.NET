using Aspose.Words.Drawing;
using Aspose.Words.Drawing.Charts;
using System;
using NUnit.Framework;

namespace Aspose.Words.Examples.CSharp.Programming_Documents.Working_with_Charts
{
    class WorkingWithChartAxis : TestDataHelper
    {
        [Test]
        public static void DefineXYAxisProperties()
        {
            //ExStart:DefineXYAxisProperties
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Insert chart
            Shape shape = builder.InsertChart(ChartType.Area, 432, 252);
            Chart chart = shape.Chart;

            // Clear demo data
            chart.Series.Clear();

            // Fill data
            chart.Series.Add("AW Series 1",
                new DateTime[]
                {
                    new DateTime(2002, 01, 01), new DateTime(2002, 06, 01), new DateTime(2002, 07, 01),
                    new DateTime(2002, 08, 01), new DateTime(2002, 09, 01)
                },
                new double[] { 640, 320, 280, 120, 150 });

            ChartAxis xAxis = chart.AxisX;
            ChartAxis yAxis = chart.AxisY;

            // Change the X axis to be category instead of date, so all the points will be put with equal interval on the X axis
            xAxis.CategoryType = AxisCategoryType.Category;

            // Define X axis properties
            xAxis.Crosses = AxisCrosses.Custom;
            xAxis.CrossesAt = 3; // measured in display units of the Y axis (hundreds)
            xAxis.ReverseOrder = true;
            xAxis.MajorTickMark = AxisTickMark.Cross;
            xAxis.MinorTickMark = AxisTickMark.Outside;
            xAxis.TickLabelOffset = 200;

            // Define Y axis properties
            yAxis.TickLabelPosition = AxisTickLabelPosition.High;
            yAxis.MajorUnit = 100;
            yAxis.MinorUnit = 50;
            yAxis.DisplayUnit.Unit = AxisBuiltInUnit.Hundreds;
            yAxis.Scaling.Minimum = new AxisBound(100);
            yAxis.Scaling.Maximum = new AxisBound(700);

            doc.Save(ArtifactsDir + "SetAxisProperties.docx");
            //ExEnd:DefineXYAxisProperties
        }

        [Test]
        public static void SetDateTimeValuesToAxis()
        {
            //ExStart:SetDateTimeValuesToAxis
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Insert chart
            Shape shape = builder.InsertChart(ChartType.Column, 432, 252);
            Chart chart = shape.Chart;

            // Clear demo data
            chart.Series.Clear();

            // Fill data
            chart.Series.Add("AW Series 1",
                new[]
                {
                    new DateTime(2017, 11, 06), new DateTime(2017, 11, 09), new DateTime(2017, 11, 15),
                    new DateTime(2017, 11, 21), new DateTime(2017, 11, 25), new DateTime(2017, 11, 29)
                },
                new double[] { 1.2, 0.3, 2.1, 2.9, 4.2, 5.3 });

            // Set X axis bounds
            ChartAxis xAxis = chart.AxisX;
            xAxis.Scaling.Minimum = new AxisBound(new DateTime(2017, 11, 05).ToOADate());
            xAxis.Scaling.Maximum = new AxisBound(new DateTime(2017, 12, 03).ToOADate());

            // Set major units to a week and minor units to a day
            xAxis.MajorUnit = 7;
            xAxis.MinorUnit = 1;
            xAxis.MajorTickMark = AxisTickMark.Cross;
            xAxis.MinorTickMark = AxisTickMark.Outside;

            doc.Save(ArtifactsDir + "SetDateTimeValuesToAxis.docx");
            //ExEnd:SetDateTimeValuesToAxis
        }

        [Test]
        public static void SetNumberFormatForAxis()
        {
            //ExStart:SetNumberFormatForAxis
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Insert chart
            Shape shape = builder.InsertChart(ChartType.Column, 432, 252);
            Chart chart = shape.Chart;

            // Clear demo data
            chart.Series.Clear();

            // Fill data
            chart.Series.Add("AW Series 1",
                new string[] { "Item 1", "Item 2", "Item 3", "Item 4", "Item 5" },
                new double[] { 1900000, 850000, 2100000, 600000, 1500000 });

            // Set number format
            chart.AxisY.NumberFormat.FormatCode = "#,##0";

            doc.Save(ArtifactsDir + "FormatAxisNumber.docx");
            //ExEnd:SetNumberFormatForAxis
        }

        [Test]
        public static void SetBoundsOfAxis()
        {
            //ExStart:SetboundsOfAxis
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Insert chart
            Shape shape = builder.InsertChart(ChartType.Column, 432, 252);
            Chart chart = shape.Chart;

            // Clear demo data
            chart.Series.Clear();

            // Fill data
            chart.Series.Add("AW Series 1",
                new string[] { "Item 1", "Item 2", "Item 3", "Item 4", "Item 5" },
                new double[] { 1.2, 0.3, 2.1, 2.9, 4.2 });

            chart.AxisY.Scaling.Minimum = new AxisBound(0);
            chart.AxisY.Scaling.Maximum = new AxisBound(6);

            doc.Save(ArtifactsDir + "SetBoundsOfAxis.docx");
            // ExEnd:SetboundsOfAxis
        }

        [Test]
        public static void SetIntervalUnitBetweenLabelsOnAxis()
        {
            //ExStart:SetIntervalUnitBetweenLabelsOnAxis
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Insert chart
            Shape shape = builder.InsertChart(ChartType.Column, 432, 252);
            Chart chart = shape.Chart;

            // Clear demo data
            chart.Series.Clear();

            // Fill data
            chart.Series.Add("AW Series 1",
                new string[] { "Item 1", "Item 2", "Item 3", "Item 4", "Item 5" },
                new double[] { 1.2, 0.3, 2.1, 2.9, 4.2 });

            chart.AxisX.TickLabelSpacing = 2;

            doc.Save(ArtifactsDir + "SetIntervalUnitBetweenLabelsOnAxis.docx");
            //ExEnd:SetIntervalUnitBetweenLabelsOnAxis
        }

        [Test]
        public static void HideChartAxis()
        {
            //ExStart:HideChartAxis
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Insert chart
            Shape shape = builder.InsertChart(ChartType.Column, 432, 252);
            Chart chart = shape.Chart;
            
            // Clear demo data
            chart.Series.Clear();
            
            // Fill data
            chart.Series.Add("AW Series 1",
                new string[] { "Item 1", "Item 2", "Item 3", "Item 4", "Item 5" },
                new double[] { 1.2, 0.3, 2.1, 2.9, 4.2 });
            
            // Hide the Y axis
            chart.AxisY.Hidden = true;

            doc.Save(ArtifactsDir + "HideChartAxis.docx");
            //ExEnd:HideChartAxis
        }

        [Test]
        public static void TickMultiLineLabelAlignment()
        {
            //ExStart:TickMultiLineLabelAlignment
            Document doc = new Document(ChartsDir + "Document.docx");

            Shape shape = (Shape) doc.GetChild(NodeType.Shape, 0, true);
            ChartAxis axis = shape.Chart.AxisX;
            // This property has effect only for multi-line labels
            axis.TickLabelAlignment = ParagraphAlignment.Right;

            doc.Save(ArtifactsDir + "Document.docx");
            //ExEnd:TickMultiLineLabelAlignment
        }
    }
}