using Aspose.Words.Drawing;
using System;
using System.IO;

namespace Aspose.Words.Examples.CSharp.Programming_Documents.Working_with_Signature
{
    class SigningSignatureLine : TestDataHelper
    {
        public static void Run()
        {
            // The path to the documents directory.
            string dataDir = RunExamples.GetDataDir_WorkingWithSignature();

            if (!File.Exists(dataDir + "signature.pfx"))
            {
                Console.WriteLine("Certificate file does not exist.");
                return;
            }

            SimpleDocumentSigning();
            SigningEncryptedDocument();
            CreatingAndSigningNewSignatureLine();
            SigningExistingSignatureLine();
            SetSignatureProviderId();
            CreateNewSignatureLineAndSetProviderId();
        }

        public static void SimpleDocumentSigning()
        {
            //ExStart:SimpleDocumentSigning
            CertificateHolder certHolder = CertificateHolder.Create(SignatureDir + "signature.pfx", "signature");
            DigitalSignatureUtil.Sign(SignatureDir + "Document.Signed.docx", ArtifactsDir + "Document.Signed.docx",
                certHolder);
            //ExEnd:SimpleDocumentSigning
        }

        public static void SigningEncryptedDocument()
        {
            //ExStart:SigningEncryptedDocument
            SignOptions signOptions = new SignOptions();
            signOptions.DecryptionPassword = "decryptionPassword";

            CertificateHolder certHolder = CertificateHolder.Create(SignatureDir + "signature.pfx", "signature");
            DigitalSignatureUtil.Sign(SignatureDir + "Document.Signed.docx", ArtifactsDir + "Document.EncryptedDocument.docx",
                certHolder, signOptions);
            //ExEnd:SigningEncryptedDocument
        }

        public static void CreatingAndSigningNewSignatureLine()
        {
            //ExStart:CreatingAndSigningNewSignatureLine
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);
            SignatureLine signatureLine = builder.InsertSignatureLine(new SignatureLineOptions()).SignatureLine;
            
            doc.Save(ArtifactsDir + "Document.NewSignatureLine.docx");

            SignOptions signOptions = new SignOptions();
            signOptions.SignatureLineId = signatureLine.Id;
            signOptions.SignatureLineImage = File.ReadAllBytes(SignatureDir + "SignatureImage.emf");

            CertificateHolder certHolder = CertificateHolder.Create(SignatureDir + "signature.pfx", "signature");
            DigitalSignatureUtil.Sign(SignatureDir + "Document.NewSignatureLine.docx",
                ArtifactsDir + "Document.NewSignatureLine.docx.docx", certHolder, signOptions);
            //ExEnd:CreatingAndSigningNewSignatureLine
        }

        public static void SigningExistingSignatureLine()
        {
            //ExStart:SigningExistingSignatureLine
            Document doc = new Document(SignatureDir + "Document.Signed.docx");
            SignatureLine signatureLine =
                ((Shape) doc.FirstSection.Body.GetChild(NodeType.Shape, 0, true)).SignatureLine;

            SignOptions signOptions = new SignOptions();
            signOptions.SignatureLineId = signatureLine.Id;
            signOptions.SignatureLineImage = File.ReadAllBytes(SignatureDir + "SignatureImage.emf");

            CertificateHolder certHolder = CertificateHolder.Create(SignatureDir + "signature.pfx", "signature");
            DigitalSignatureUtil.Sign(SignatureDir + "Document.Signed.docx",
                ArtifactsDir + "Document.Signed.ExistingSignatureLine.docx", certHolder, signOptions);
            //ExEnd:SigningExistingSignatureLine
        }

        public static void SetSignatureProviderId()
        {
            //ExStart:SetSignatureProviderID
            Document doc = new Document(SignatureDir + "Document.Signed.docx");
            SignatureLine signatureLine =
                ((Shape) doc.FirstSection.Body.GetChild(NodeType.Shape, 0, true)).SignatureLine;

            // Set signature and signature line provider ID
            SignOptions signOptions = new SignOptions();
            signOptions.ProviderId = signatureLine.ProviderId;
            signOptions.SignatureLineId = signatureLine.Id;

            CertificateHolder certHolder = CertificateHolder.Create(SignatureDir + "signature.pfx", "signature");
            DigitalSignatureUtil.Sign(SignatureDir + "Document.Signed.docx", ArtifactsDir + "Document.Signed.docx",
                certHolder, signOptions);
            //ExEnd:SetSignatureProviderID
        }

        public static void CreateNewSignatureLineAndSetProviderId()
        {
            //ExStart:CreateNewSignatureLineAndSetProviderID
            Document doc = new Document(SignatureDir + "Document.Signed.docx");
            DocumentBuilder builder = new DocumentBuilder(doc);
            SignatureLine signatureLine = builder.InsertSignatureLine(new SignatureLineOptions()).SignatureLine;
            signatureLine.ProviderId = new Guid("{F5AC7D23-DA04-45F5-ABCB-38CE7A982553}");
            
            doc.Save(ArtifactsDir + "Document.Signed.docx");

            SignOptions signOptions = new SignOptions();
            signOptions.SignatureLineId = signatureLine.Id;
            signOptions.ProviderId = signatureLine.ProviderId;

            CertificateHolder certHolder = CertificateHolder.Create(SignatureDir + "signature.pfx", "signature");
            DigitalSignatureUtil.Sign(ArtifactsDir + "Document.Signed.docx", ArtifactsDir + "Document.Signed_out.docx",
                certHolder, signOptions);
            //ExEnd:CreateNewSignatureLineAndSetProviderID
        }
    }
}