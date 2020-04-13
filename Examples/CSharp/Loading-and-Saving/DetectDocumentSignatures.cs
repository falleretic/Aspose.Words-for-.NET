using System.IO;
using System;

namespace Aspose.Words.Examples.CSharp.Loading_Saving
{
    class DetectDocumentSignatures : TestDataHelper
    {
        public static void Run()
        {
            //ExStart:DetectDocumentSignatures
            string filePath = LoadingSavingDir + "Document.Signed.docx";

            FileFormatInfo info = FileFormatUtil.DetectFileFormat(filePath);
            if (info.HasDigitalSignature)
            {
                Console.WriteLine(
                    "Document {0} has digital signatures, they will be lost if you open/save this document with Aspose.Words.",
                    Path.GetFileName(filePath));
            }
            //ExEnd:DetectDocumentSignatures            
        }
    }
}