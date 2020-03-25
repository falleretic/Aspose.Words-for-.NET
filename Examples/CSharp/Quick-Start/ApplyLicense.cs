using System;

namespace Aspose.Words.Examples.CSharp.Quick_Start
{
    class ApplyLicense
    {
        public static void Run()
        {
            License license = new License();

            // This line attempts to set a license from several locations relative to the executable and Aspose.Words.dll.
            // You can also use the additional overload to load a license from a stream, this is useful for instance when the 
            // License is stored as an embedded resource 
            try
            {
                license.SetLicense(@"Aspose.Total.NET.lic");
                Console.WriteLine("License set successfully.");
            }

            catch (Exception e)
            {
                // We do not ship any license with this example, visit the Aspose site to obtain either a temporary or permanent license. 
                Console.WriteLine("\nThere was an error setting the license: " + e.Message);
            }
        }
    }
}