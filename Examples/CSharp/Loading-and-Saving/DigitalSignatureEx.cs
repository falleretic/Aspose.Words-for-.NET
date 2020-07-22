﻿using System;
using NUnit.Framework;

namespace Aspose.Words.Examples.CSharp
{
    class DigitalSignatureEx : TestDataHelper
    {
        [Test]
        public static void AccessAndVerifySignature()
        {
            //ExStart:AccessAndVerifySignature
            Document doc = new Document(LoadingSavingDir + "Digitally signed.docx");

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