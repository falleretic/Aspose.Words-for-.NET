using System.Data;
using NUnit.Framework;

namespace Aspose.Words.Examples.CSharp
{
    class XMLMailMerge : TestDataHelper
    {
        [Test]
        public static void Run()
        {
            //ExStart:XMLMailMerge
            DataSet customersDs = new DataSet();
            customersDs.ReadXml(MyDir + "Mail merge data - Customers.xml");

            Document doc = new Document(MyDir + "Mail merge destinations - Registration complete.docx");
            // Execute mail merge to fill the template with data from XML using DataTable
            doc.MailMerge.Execute(customersDs.Tables["Customer"]);

            doc.Save(ArtifactsDir + "XMLMailMerge.docx");
            //ExEnd:XMLMailMerge
        }
    }
}