using Aspose.Words.Reporting;
using NUnit.Framework;

namespace Aspose.Words.Examples.CSharp.LINQ_Reporting_Engine
{
    internal class BuildOptions : TestDataHelper
    {
        [Test]
        public static void RemoveEmptyParagraphs()
        {
            //ExStart:RemoveEmptyParagraphs
            Document doc = new Document(LinqDir + "Reporting engine template - Empty paragraphs.docx");
            ReportingEngine engine = new ReportingEngine { Options = ReportBuildOptions.RemoveEmptyParagraphs };
            
            engine.BuildReport(doc, Helpers.Common.GetManagers(), "managers");

            doc.Save(ArtifactsDir + "ReportingEngine.RemoveEmptyParagraphs.docx");
            //ExEnd:RemoveEmptyParagraphs
        }
    }
}