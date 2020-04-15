﻿using Aspose.Words.Reporting;
using NUnit.Framework;

namespace Aspose.Words.Examples.CSharp.LINQ
{
    class PieChart : TestDataHelper
    {
        [Test]
        public static void Run()
        {
            //ExStart:PieChart
            Document doc = new Document(LinqDir + "PieChart.docx");

            // Create a Reporting Engine
            ReportingEngine engine = new ReportingEngine();
            // Execute the build report
            engine.BuildReport(doc, Common.GetManagers(), "managers");

            doc.Save(ArtifactsDir + "PieChart.docx");
            //ExEnd:PieChart
        }
    }
}