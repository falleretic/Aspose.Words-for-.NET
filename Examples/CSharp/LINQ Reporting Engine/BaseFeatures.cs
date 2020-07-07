using Aspose.Words.Examples.CSharp.LINQ_Reporting_Engine.Helpers.Data_Source_Objects;
using Aspose.Words.Reporting;
using NUnit.Framework;

namespace Aspose.Words.Examples.CSharp.LINQ_Reporting_Engine
{
    internal class BaseFeatures : TestDataHelper
    {
        [Test]
        public static void HelloWorld()
        {
            //ExStart:HelloWorld
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);
            
            builder.Write("<<[sender.Name]>> says: <<[sender.Message]>>");

            Sender sender = new Sender { Name = "LINQ Reporting Engine", Message = "Hello World" };

            ReportingEngine engine = new ReportingEngine();
            engine.BuildReport(doc, sender, "sender");

            doc.Save(ArtifactsDir + "ReportingEngine.HelloWorld.docx");
            //ExEnd:HelloWorld
        }

        [Test]
        public static void SingleRow()
        {
            //ExStart:SingleRow
            Document doc = new Document(LinqDir + "Reporting engine template - Single row.docx");

            ReportingEngine engine = new ReportingEngine();
            engine.BuildReport(doc, Helpers.Common.GetManager(), "manager");

            doc.Save(ArtifactsDir + "ReportingEngine.SingleRow.docx");
            //ExEnd:SingleRow
        }

        [Test]
        public static void CommonMasterDetail()
        {
            //ExStart:CommonMasterDetail
            Document doc = new Document(LinqDir + "Reporting engine template - Common master detail.docx");

            ReportingEngine engine = new ReportingEngine();
            engine.BuildReport(doc, Helpers.Common.GetManagers(), "managers");

            doc.Save(ArtifactsDir + "ReportingEngine.CommonMasterDetail.docx");
            //ExEnd:CommonMasterDetail
        }

        [Test]
        public static void ConditionalBlocks()
        {
            //ExStart:ConditionalBlocks
            Document doc = new Document(LinqDir + "Reporting engine template - Conditional block.docx");

            ReportingEngine engine = new ReportingEngine();
            engine.BuildReport(doc, Helpers.Common.GetClients(), "clients");

            doc.Save(ArtifactsDir + "ReportingEngine.ConditionalBlock.docx");
            //ExEnd:ConditionalBlocks
        }

        [Test]
        public static void SettingBackgroundColor()
        {
            //ExStart:SettingBackgroundColor
            Document doc = new Document(LinqDir + "Reporting engine template - Background color.docx");

            ReportingEngine engine = new ReportingEngine();
            engine.BuildReport(doc, new object());

            doc.Save(ArtifactsDir + "ReportingEngine.SettingBackgroundColor.docx");
            //ExEnd:SettingBackgroundColor
        }
    }
}