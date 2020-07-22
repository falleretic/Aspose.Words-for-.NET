using NUnit.Framework;

namespace Aspose.Words.Examples.CSharp.DocumentEx
{
    class ProtectDocument : TestDataHelper
    {
        /// <summary>
        /// Shows how to protect document.
        /// </summary>
        [Test]
        public static void Protect()
        {
            //ExStart:ProtectDocument
            Document doc = new Document(DocumentDir + "Document.docx");
            doc.Protect(ProtectionType.AllowOnlyFormFields, "password");
            //ExEnd:ProtectDocument
        }

        /// <summary>
        /// Shows how to unprotect document.
        /// </summary>
        [Test]
        public static void UnProtect()
        {
            // ExStart:UnProtectDocument
            Document doc = new Document(DocumentDir + "Document.docx");
            doc.Unprotect();
            // ExEnd:UnProtectDocument
        }

        /// <summary>
        /// Shows how to get protection type.
        /// </summary>
        [Test]
        public static void GetProtectionType()
        {
            //ExStart:GetProtectionType
            Document doc = new Document(DocumentDir + "Document.docx");
            ProtectionType protectionType = doc.ProtectionType;
            //ExEnd:GetProtectionType
        }
    }
}