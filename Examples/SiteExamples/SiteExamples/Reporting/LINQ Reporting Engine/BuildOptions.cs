using Aspose.Words;
using Aspose.Words.Reporting;
using NUnit.Framework;

namespace SiteExamples.Reporting.LINQ_Reporting_Engine
{
    internal class BuildOptions : SiteExamplesBase
    {
        [Test]
        public static void RemoveEmptyParagraphs()
        {
            //ExStart:RemoveEmptyParagraphs
            Document doc = new Document(MyDir + "Reporting engine template - Remove empty paragraphs.docx");
            ReportingEngine engine = new ReportingEngine { Options = ReportBuildOptions.RemoveEmptyParagraphs };
            
            engine.BuildReport(doc, Helpers.Common.GetManagers(), "Managers");

            doc.Save(ArtifactsDir + "ReportingEngine.RemoveEmptyParagraphs.docx");
            //ExEnd:RemoveEmptyParagraphs
        }
    }
}