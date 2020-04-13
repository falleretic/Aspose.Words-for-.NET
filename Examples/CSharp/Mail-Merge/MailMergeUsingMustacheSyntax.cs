using System.Data;

namespace Aspose.Words.Examples.CSharp.Mail_Merge
{
    class MailMergeUsingMustacheSyntax : TestDataHelper
    {
        public static void Run()
        {
            MustacheSyntax();
            UseIfElseMustacheSyntax();
        }

        public static void MustacheSyntax()
        {
            //ExStart:MailMergeUsingMustacheSyntax
            DataSet ds = new DataSet();
            ds.ReadXml(MailMergeDir + "Vendors.xml");

            // Open a template document
            Document doc = new Document(MailMergeDir + "VendorTemplate.doc");

            doc.MailMerge.UseNonMergeFields = true;

            // Execute mail merge to fill the template with data from XML using DataSet
            doc.MailMerge.ExecuteWithRegions(ds);
            
            doc.Save(ArtifactsDir + "MailMergeUsingMustacheSyntax.docx");
            //ExEnd:MailMergeUsingMustacheSyntax
        }

        public static void UseIfElseMustacheSyntax()
        {
            //ExStart:UseOfifelseMustacheSyntax
            Document doc = new Document(MailMergeDir + "UseOfifelseMustacheSyntax.docx");

            doc.MailMerge.UseNonMergeFields = true;
            doc.MailMerge.Execute(new string[] { "GENDER" }, new object[] { "MALE" });

            doc.Save(ArtifactsDir + "UseIfElseMustacheSyntax.docx");
            //ExEnd:UseOfifelseMustacheSyntax
        }
    }
}