using Aspose.Words;
using NUnit.Framework;

namespace SiteExamples.Programming_with_Documents.Document_Content
{
    class DocumentProtection : SiteExamplesBase
    {
        [Test]
        public static void Protect()
        {
            //ExStart:ProtectDocument
            Document doc = new Document(MyDir + "Document.docx");
            doc.Protect(ProtectionType.AllowOnlyFormFields, "password");
            //ExEnd:ProtectDocument
        }

        [Test]
        public static void Unprotect()
        {
            // ExStart:UnprotectDocument
            Document doc = new Document(MyDir + "Document.docx");
            doc.Unprotect();
            // ExEnd:UnprotectDocument
        }

        [Test]
        public static void GetProtectionType()
        {
            //ExStart:GetProtectionType
            Document doc = new Document(MyDir + "Document.docx");
            ProtectionType protectionType = doc.ProtectionType;
            //ExEnd:GetProtectionType
        }
    }
}