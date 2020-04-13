using Aspose.Words.Reporting;

namespace Aspose.Words.Examples.CSharp.LINQ
{
    class InTableWithFilteringGroupingSorting : TestDataHelper
    {
        public static void Run()
        {
            //ExStart:InTableWithFilteringGroupingSorting
            Document doc = new Document(LinqDir + "InTableWithFilteringGroupingSorting.doc");

            // Create a Reporting Engine
            ReportingEngine engine = new ReportingEngine();
            // Execute the build report
            engine.BuildReport(doc, Common.GetContracts(), "contracts");

            doc.Save(ArtifactsDir + "InTableWithFilteringGroupingSorting.docx");
            //ExEnd:InTableWithFilteringGroupingSorting
        }
    }
}