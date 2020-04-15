using Aspose.Words.Reporting;
using NUnit.Framework;

namespace Aspose.Words.Examples.CSharp.LINQ
{
    class HelloWorld : TestDataHelper
    {
        [Test]
        public static void Run()
        {
            //ExStart:HelloWorld
            Document doc = new Document(LinqDir + "HelloWorld.doc");

            // Create an instance of sender class to set it's properties
            Sender sender = new Sender { Name = "LINQ Reporting Engine", Message = "Hello World" };

            // Create a Reporting Engine
            ReportingEngine engine = new ReportingEngine();
            // Execute the build report
            engine.BuildReport(doc, sender, "sender");

            doc.Save(ArtifactsDir + "HelloWorld.doc");
            //ExEnd:HelloWorld
        }
    }
}