using Aspose.Words.Reporting;

namespace Aspose.Words.Examples.CSharp.LINQ
{
    class InTableAlternateContent : TestDataHelper
    {
        public static void Run()
        {
            //ExStart:InTableAlternateContent
            Document doc = new Document(LinqDir + "InTableAlternateContent.doc");

            // Create a Reporting Engine
            ReportingEngine engine = new ReportingEngine();
            // Execute the build report
            engine.BuildReport(doc, Common.GetContracts(), "contracts");

            doc.Save(ArtifactsDir + "InTableAlternateContent.docx");
            //ExEnd:InTableAlternateContent
        }
    }
}