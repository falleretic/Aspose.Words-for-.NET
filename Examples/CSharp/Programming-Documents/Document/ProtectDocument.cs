namespace Aspose.Words.Examples.CSharp.Programming_Documents.Working_With_Document
{
    class ProtectDocument : TestDataHelper
    {
        public static void Run()
        {
            Protect();
            UnProtect();
            GetProtectionType();
        }

        /// <summary>
        /// Shows how to protect document.
        /// </summary>      
        private static void Protect()
        {
            //ExStart:ProtectDocument
            Document doc = new Document(DocumentDir + "ProtectDocument.doc");
            doc.Protect(ProtectionType.AllowOnlyFormFields, "password");
            //ExEnd:ProtectDocument
        }

        /// <summary>
        /// Shows how to unprotect document.
        /// </summary>      
        private static void UnProtect()
        {
            // ExStart:UnProtectDocument
            Document doc = new Document(DocumentDir + "ProtectDocument.doc");
            doc.Unprotect();
            // ExEnd:UnProtectDocument
        }

        /// <summary>
        /// Shows how to get protection type.
        /// </summary>        
        private static void GetProtectionType()
        {
            //ExStart:GetProtectionType
            Document doc = new Document(DocumentDir + "ProtectDocument.doc");
            ProtectionType protectionType = doc.ProtectionType;
            //ExEnd:GetProtectionType
        }
    }
}