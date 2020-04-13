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
            BookmarksDir = new Uri(new Uri(MainDataDir), @"Programming-Documents/Bookmarks/").LocalPath;
            ChartsDir = new Uri(new Uri(MainDataDir), @"Programming-Documents/Charts/").LocalPath;
            CommentsDir = new Uri(new Uri(MainDataDir), @"Programming-Documents/Comments/").LocalPath;
            DocumentDir = new Uri(new Uri(MainDataDir), @"Programming-Documents/Document/").LocalPath;
            FieldsDir = new Uri(new Uri(MainDataDir), @"Programming-Documents/Fields/").LocalPath;
            FindReplaceDir = new Uri(new Uri(MainDataDir), @"Programming-Documents/Find-Replace/").LocalPath;
            HyperlinkDir = new Uri(new Uri(MainDataDir), @"Programming-Documents/Hyperlink/").LocalPath;
            ImagesDir = new Uri(new Uri(MainDataDir), @"Programming-Documents/Images/").LocalPath;
            JoiningAppendingDir = new Uri(new Uri(MainDataDir), @"Programming-Documents/Joining-Appending/").LocalPath;
            NodeDir = new Uri(new Uri(MainDataDir), @"Programming-Documents/Node/").LocalPath;
            RangeDir = new Uri(new Uri(MainDataDir), @"Programming-Documents/Ranges/").LocalPath;
            SectionsDir = new Uri(new Uri(MainDataDir), @"Programming-Documents/Sections/").LocalPath;
            ShapesDir = new Uri(new Uri(MainDataDir), @"Programming-Documents/Shapes/").LocalPath;
            SignatureDir = new Uri(new Uri(MainDataDir), @"Programming-Documents/Signature/").LocalPath;
            SdtDir = new Uri(new Uri(MainDataDir), @"Programming-Documents/StructuredDocumentTag/").LocalPath;
            StyleDir = new Uri(new Uri(MainDataDir), @"Programming-Documents/Styles/").LocalPath;
            ThemeDir = new Uri(new Uri(MainDataDir), @"Programming-Documents/Theme/").LocalPath;
            WebExtensionsDir = new Uri(new Uri(MainDataDir), @"Programming-Documents/WebExtensions/").LocalPath;
            LoadingSavingDir = new Uri(new Uri(MainDataDir), @"Loading-and-Saving/").LocalPath;
            DatabaseDir = new Uri(new Uri(MainDataDir), @"Database/").LocalPath;
            TablesDir = new Uri(new Uri(MainDataDir), @"Programming-Documents/Tables/").LocalPath;
            LinqDir = new Uri(new Uri(MainDataDir), @"LINQ/").LocalPath;
            MailMergeDir = new Uri(new Uri(MainDataDir), @"Mail-Merge/").LocalPath;
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
        /// Gets the path to the license used by the code examples.
        /// </summary>
        internal static string BookmarksDir { get; }

        /// <summary>
        /// Gets the path to the license used by the code examples.
        /// </summary>
        internal static string ChartsDir { get; }

        /// <summary>
        /// Gets the path to the license used by the code examples.
        /// </summary>
        internal static string CommentsDir { get; }

        /// <summary>
        /// Gets the path to the license used by the code examples.
        /// </summary>
        internal static string DocumentDir { get; }

        /// <summary>
        /// Gets the path to the license used by the code examples.
        /// </summary>
        internal static string FieldsDir { get; }

        /// <summary>
        /// Gets the path to the license used by the code examples.
        /// </summary>
        internal static string FindReplaceDir { get; }

        /// <summary>
        /// Gets the path to the license used by the code examples.
        /// </summary>
        internal static string HyperlinkDir { get; }

        /// <summary>
        /// Gets the path to the license used by the code examples.
        /// </summary>
        internal static string ImagesDir { get; }

        /// <summary>
        /// Gets the path to the license used by the code examples.
        /// </summary>
        internal static string JoiningAppendingDir { get; }

        /// <summary>
        /// Gets the path to the license used by the code examples.
        /// </summary>
        internal static string NodeDir { get; }

        /// <summary>
        /// Gets the path to the license used by the code examples.
        /// </summary>
        internal static string RangeDir { get; }

        /// <summary>
        /// Gets the path to the license used by the code examples.
        /// </summary>
        internal static string SectionsDir { get; }

        /// <summary>
        /// Gets the path to the license used by the code examples.
        /// </summary>
        internal static string ShapesDir { get; }

        /// <summary>
        /// Gets the path to the license used by the code examples.
        /// </summary>
        internal static string SignatureDir { get; }

        /// <summary>
        /// Gets the path to the license used by the code examples.
        /// </summary>
        internal static string SdtDir { get; }

        /// <summary>
        /// Gets the path to the license used by the code examples.
        /// </summary>
        internal static string StyleDir { get; }

        /// <summary>
        /// Gets the path to the license used by the code examples.
        /// </summary>
        internal static string ThemeDir { get; }

        /// <summary>
        /// Gets the path to the license used by the code examples.
        /// </summary>
        internal static string WebExtensionsDir { get; }

        /// <summary>
        /// Gets the path to the license used by the code examples.
        /// </summary>
        internal static string TablesDir { get; }

        /// <summary>
        /// Gets the path to the license used by the code examples.
        /// </summary>
        internal static string LoadingSavingDir { get; }

        /// <summary>
        /// Gets the path to the license used by the code examples.
        /// </summary>
        internal static string DatabaseDir { get; }

        /// <summary>
        /// Gets the path to the license used by the code examples.
        /// </summary>
        internal static string LinqDir { get; }

        /// <summary>
        /// Gets the path to the license used by the code examples.
        /// </summary>
        internal static string MailMergeDir { get; }

        /// <summary>
        /// Gets the path to the documents used by the code examples. Ends with a back slash.
        /// </summary>
        internal static string ArtifactsDir { get; }
    }
}