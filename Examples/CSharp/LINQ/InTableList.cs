using Aspose.Words.Reporting;

namespace Aspose.Words.Examples.CSharp.LINQ
{
    class InTableList : TestDataHelper
    {
        public static void Run()
        {
            //ExStart:InTableList
            Document doc = new Document(LinqDir + "InTableList.doc");

            // Create a Reporting Engine
            ReportingEngine engine = new ReportingEngine();
            // Execute the build report
            engine.BuildReport(doc, Common.GetManagers(), "managers");

            doc.Save(ArtifactsDir + "InTableList.docx");
            //ExEnd:InTableList
        }
    }
}