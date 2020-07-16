using System;
using System.IO;
using NUnit.Framework;

namespace Aspose.Words.Examples.CSharp
{
    class ApplyLicenseFromStream
    {
        [Test]
        public static void Run()
        {
            //ExStart:ApplyLicenseFromStream
            License license = new License();

            try
            {
                // Initializes a license from a stream 
                MemoryStream stream = new MemoryStream(File.ReadAllBytes(@"Aspose.Words.lic"));
                license.SetLicense(stream);
                Console.WriteLine("License set successfully.");
            }
            catch (Exception e)
            {
                // We do not ship any license with this example, visit the Aspose site to obtain either a temporary or permanent license. 
                Console.WriteLine("\nThere was an error setting the license: " + e.Message);
            }
            //ExEnd:ApplyLicenseFromStream
        }
    }
}