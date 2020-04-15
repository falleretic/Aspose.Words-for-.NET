using Aspose.Words.Reporting;
using NUnit.Framework;

namespace Aspose.Words.Examples.CSharp.LINQ
{
    class InTableList : TestDataHelper
    {
        [Test]
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