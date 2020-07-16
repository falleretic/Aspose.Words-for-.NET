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
            customersDs.ReadXml(MailMergeDir + "Customers.xml");

            Document doc = new Document(MailMergeDir + "TestFile XML.doc");
            // Execute mail merge to fill the template with data from XML using DataTable
            doc.MailMerge.Execute(customersDs.Tables["Customer"]);

            doc.Save(ArtifactsDir + "XMLMailMerge.docx");
            //ExEnd:XMLMailMerge
        }
    }
}