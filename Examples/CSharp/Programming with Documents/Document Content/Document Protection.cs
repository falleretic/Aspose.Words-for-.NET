using NUnit.Framework;

namespace Aspose.Words.Examples.CSharp.Programming_with_Documents.Document_Content
{
    class DocumentProtection : TestDataHelper
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
        public static void UnProtect()
        {
            // ExStart:UnProtectDocument
            Document doc = new Document(MyDir + "Document.docx");
            doc.Unprotect();
            // ExEnd:UnProtectDocument
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