using System;

namespace Aspose.Words.Examples.CSharp.Loading_Saving
{
    class AccessAndVerifySignature : TestDataHelper
    {
        public static void Run()
        {
            //ExStart:AccessAndVerifySignature
            Document doc = new Document(LoadingSavingDir + "Test File (doc).doc");

            foreach (DigitalSignature signature in doc.DigitalSignatures)
            {
                Console.WriteLine("*** Signature Found ***");
                Console.WriteLine("Is valid: " + signature.IsValid);
                // This property is available in MS Word documents only
                Console.WriteLine("Reason for signing: " + signature.Comments); 
                Console.WriteLine("Time of signing: " + signature.SignTime);
                Console.WriteLine("Subject name: " + signature.CertificateHolder.Certificate.SubjectName.Name);
                Console.WriteLine("Issuer name: " + signature.CertificateHolder.Certificate.IssuerName.Name);
                Console.WriteLine();
            }
            //ExEnd:AccessAndVerifySignature
        }
    }
}