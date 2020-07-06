using Aspose.Words.Reporting;
using System.Collections.Generic;
using NUnit.Framework;

namespace Aspose.Words.Examples.CSharp.LINQ
{
    class ChartSeries : TestDataHelper
    {
        [Test]
        public static void SetChartSeriesNameDynamically()
        {
            // ExStart:SetChartSeriesNameDynamically
            List<PointData> data = new List<PointData>()
            {
                new PointData { Time = "12:00:00 AM", Flow = 10, Rainfall = 2 },
                new PointData { Time = "01:00:00 AM", Flow = 15, Rainfall = 4 },
                new PointData { Time = "02:00:00 AM", Flow = 23, Rainfall = 7 }
            };

                        List<string> seriesNames = new List<string>
            {
                "Flow",
                "Rainfall"
            };

            Document doc = new Document(LinqDir + "ChartTemplate.docx");

            ReportingEngine engine = new ReportingEngine();
            engine.BuildReport(doc, new object[] { data, seriesNames }, new string[] { "data", "seriesNames" });

            doc.Save(ArtifactsDir + "ChartTemplate.docx");
            // ExEnd:SetChartSeriesNameDynamically
        }

    }
    // ExStart:PointDataClass
    public class PointData
    {
        public string Time { get; set; }
        public int Flow { get; set; }
        public int Rainfall { get; set; }
    }
    // ExEnd:PointDataClass
}
