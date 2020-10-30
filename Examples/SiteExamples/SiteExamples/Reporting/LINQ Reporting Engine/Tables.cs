using Aspose.Words;
using Aspose.Words.Reporting;
using NUnit.Framework;

namespace SiteExamples.Reporting.LINQ_Reporting_Engine
{
    internal class Tables : SiteExamplesBase
    {
        [Test]
        public static void InTableAlternateContent()
        {
            //ExStart:InTableAlternateContent
            Document doc = new Document(MyDir + "Reporting engine template - Total.docx");

            ReportingEngine engine = new ReportingEngine();
            engine.BuildReport(doc, Helpers.Common.GetContracts(), "Contracts");

            doc.Save(ArtifactsDir + "ReportingEngine.InTableAlternateContent.docx");
            //ExEnd:InTableAlternateContent
        }

        [Test]
        public static void InTableMasterDetail()
        {
            //ExStart:InTableMasterDetail
            Document doc = new Document(MyDir + "Reporting engine template - Nested data table.docx");

            ReportingEngine engine = new ReportingEngine();
            engine.BuildReport(doc, Helpers.Common.GetManagers(), "Managers");

            doc.Save(ArtifactsDir + "ReportingEngine.InTableMasterDetail.docx");
            //ExEnd:InTableMasterDetail
        }

        [Test]
        public static void InTableWithFilteringGroupingSorting()
        {
            //ExStart:InTableWithFilteringGroupingSorting
            Document doc = new Document(MyDir + "Reporting engine template - Table with filtering.docx");

            ReportingEngine engine = new ReportingEngine();
            engine.BuildReport(doc, Helpers.Common.GetContracts(), "contracts");

            doc.Save(ArtifactsDir + "ReportingEngine.InTableWithFilteringGroupingSorting.docx");
            //ExEnd:InTableWithFilteringGroupingSorting
        }
    }
}