using System.Data;
using NUnit.Framework;

namespace Aspose.Words.Examples.CSharp
{
    class MailMergeUsingMustacheSyntax : TestDataHelper
    {
        [Test]
        public static void MustacheSyntax()
        {
            //ExStart:MailMergeUsingMustacheSyntax
            DataSet ds = new DataSet();
            ds.ReadXml(MailMergeDir + "Mail merge data - Vendors.xml");

            // Open a template document
            Document doc = new Document(MailMergeDir + "VendorTemplate.doc");

            doc.MailMerge.UseNonMergeFields = true;

            // Execute mail merge to fill the template with data from XML using DataSet
            doc.MailMerge.ExecuteWithRegions(ds);
            
            doc.Save(ArtifactsDir + "MailMerge.UsingMustacheSyntax.docx");
            //ExEnd:MailMergeUsingMustacheSyntax
        }

        [Test]
        public static void UseIfElseMustacheSyntax()
        {
            //ExStart:UseOfifelseMustacheSyntax
            Document doc = new Document(MailMergeDir + "Mail merge destinations - Mustache syntax.docx");

            doc.MailMerge.UseNonMergeFields = true;
            doc.MailMerge.Execute(new string[] { "GENDER" }, new object[] { "MALE" });

            doc.Save(ArtifactsDir + "MailMerge.IfElseMustacheSyntax.docx");
            //ExEnd:UseOfifelseMustacheSyntax
        }
    }
}