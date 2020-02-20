using System;
using System.IO;

namespace Aspose.Words.Examples.CSharp
{
    internal class TestDataHelper
    {
        static TestDataHelper()
        {
            MainDataDir = GetDataDir_Data();
            ArtifactsDir = new Uri(new Uri(MainDataDir), @"Artifacts/").LocalPath;
            QuickStartDir = new Uri(new Uri(MainDataDir), @"Quick-Start/").LocalPath;
        }

        private static string GetDataDir_Data()
        {
            DirectoryInfo parent = Directory.GetParent(Directory.GetCurrentDirectory()).Parent;
            string startDirectory = null;
            if (parent != null)
            {
                DirectoryInfo directoryInfo = parent.Parent;
                if (directoryInfo != null)
                {
                    startDirectory = directoryInfo.FullName;
                }
            }
            else
            {
                startDirectory = parent.FullName;
            }

            return Path.Combine(startDirectory, "Data\\");
        }

        /// <summary>
        /// Gets the path to the codebase directory.
        /// </summary>
        internal static string MainDataDir { get; }

        /// <summary>
        /// Gets the path to the license used by the code examples.
        /// </summary>
        internal static string QuickStartDir { get; }

        /// <summary>
        /// Gets the path to the documents used by the code examples. Ends with a back slash.
        /// </summary>
        internal static string ArtifactsDir { get; }
    }
}