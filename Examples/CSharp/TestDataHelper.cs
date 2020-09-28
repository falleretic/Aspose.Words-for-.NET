using System;
using System.Globalization;
using System.IO;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Threading;
using NUnit.Framework;

namespace Aspose.Words.Examples.CSharp
{
    internal class TestDataHelper
    {
        static TestDataHelper()
        {
            MainDataDir = GetCodeBaseDir(Assembly.GetExecutingAssembly());
            ArtifactsDir = new Uri(new Uri(MainDataDir), @"Data/Artifacts/").LocalPath;
            MyDir = new Uri(new Uri(MainDataDir), @"Data/").LocalPath;
            ImagesDir = new Uri(new Uri(MainDataDir), @"Data/Images/").LocalPath;
        }

        [OneTimeSetUp]
        public void OneTimeSetUp()
        {
            Thread.CurrentThread.CurrentCulture = CultureInfo.InvariantCulture;

            if (!Directory.Exists(ArtifactsDir))
                //Create new empty directory
                Directory.CreateDirectory(ArtifactsDir);
        }

        [SetUp]
        public void SetUp()
        {
            Console.WriteLine($"Clr: {RuntimeInformation.FrameworkDescription}\n");
        }

        [OneTimeTearDown]
        public void OneTimeTearDown()
        {
            if (Directory.Exists(ArtifactsDir))
                //Delete all dirs and files from directory
                Directory.Delete(ArtifactsDir, true);
        }

        /// <summary>
        /// Returns the code-base directory.
        /// </summary>
        internal static string GetCodeBaseDir(Assembly assembly)
        {
            // CodeBase is a full URI, such as file:///x:\blahblah.
            Uri uri = new Uri(assembly.CodeBase);
            string mainFolder = Path.GetDirectoryName(uri.LocalPath)
                ?.Substring(0, uri.LocalPath.IndexOf("CSharp", StringComparison.Ordinal));
            return mainFolder;
        }

        /// <summary>
        /// Gets the path to the codebase directory.
        /// </summary>
        internal static string MainDataDir { get; }

        /// <summary>
        /// Gets the path to the documents used by the code examples.
        /// </summary>
        internal static string MyDir { get; }

        /// <summary>
        /// Gets the path to the images used by the code examples.
        /// </summary>
        internal static string ImagesDir { get; }

        /// <summary>
        /// Gets the path to the artifacts used by the code examples.
        /// </summary>
        internal static string ArtifactsDir { get; }
    }
}