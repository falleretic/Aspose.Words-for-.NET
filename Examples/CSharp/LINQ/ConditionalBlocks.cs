using Aspose.Words.Reporting;
using NUnit.Framework;

namespace Aspose.Words.Examples.CSharp.LINQ
{
    class ConditionalBlocks : TestDataHelper
    {
        [Test]
        public static void Run()
        {
            //ExStart:InTableList
            Document doc = new Document(LinqDir + "ConditionalBlock.doc");

            // Create a Reporting Engine
            ReportingEngine engine = new ReportingEngine();
            // Execute the build report
            engine.BuildReport(doc, Common.GetClients(), "clients");

            doc.Save(ArtifactsDir + "ConditionalBlock.docx");
            //ExEnd:InTableList
        }
    }
}