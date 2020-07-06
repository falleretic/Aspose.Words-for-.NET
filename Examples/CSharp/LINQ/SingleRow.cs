using Aspose.Words.Reporting;
using NUnit.Framework;

namespace Aspose.Words.Examples.CSharp.LINQ
{
    class SingleRow : TestDataHelper
    {
        [Test]
        public static void Run()
        {
            //ExStart:SingleRow
            Document doc = new Document(LinqDir + "SingleRow.doc");

            // Create a Reporting Engine
            ReportingEngine engine = new ReportingEngine();
            // Execute the build report
            engine.BuildReport(doc, Common.GetManager(), "manager");

            doc.Save(ArtifactsDir + "SingleRow.docx");
            //ExEnd:SingleRow
        }
    }
}