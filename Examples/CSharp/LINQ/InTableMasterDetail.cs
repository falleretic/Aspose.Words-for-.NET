using Aspose.Words.Reporting;

namespace Aspose.Words.Examples.CSharp.LINQ
{
    class InTableMasterDetail : TestDataHelper
    {
        public static void Run()
        {
            //ExStart:InTableMasterDetail
            Document doc = new Document(LinqDir + "InTableMasterDetail.doc");

            // Create a Reporting Engine
            ReportingEngine engine = new ReportingEngine();
            // Execute the build report
            engine.BuildReport(doc, Common.GetManagers(), "managers");

            doc.Save(ArtifactsDir + "InTableMasterDetail.docx");
            //ExEnd:InTableMasterDetail
        }
    }
}