using System;
using System.IO;
using System.Reflection;

namespace Aspose.Words.Examples.CSharp
{
    internal class TestDataHelper
    {
        static TestDataHelper()
        {
            MainDataDir = GetCodeBaseDir(Assembly.GetExecutingAssembly());
            ArtifactsDir = new Uri(new Uri(MainDataDir), @"Data/Artifacts/").LocalPath;
            QuickStartDir = new Uri(new Uri(MainDataDir), @"Data/Quick-Start/").LocalPath;
            BookmarksDir = new Uri(new Uri(MainDataDir), @"Data/Programming-Documents/Bookmarks/").LocalPath;
            ChartsDir = new Uri(new Uri(MainDataDir), @"Data/Programming-Documents/Charts/").LocalPath;
            CommentsDir = new Uri(new Uri(MainDataDir), @"Data/Programming-Documents/Comments/").LocalPath;
            DocumentDir = new Uri(new Uri(MainDataDir), @"Data/Programming-Documents/Document/").LocalPath;
            FieldsDir = new Uri(new Uri(MainDataDir), @"Data/Programming-Documents/Fields/").LocalPath;
            FindReplaceDir = new Uri(new Uri(MainDataDir), @"Data/Programming-Documents/Find-Replace/").LocalPath;
            HyperlinkDir = new Uri(new Uri(MainDataDir), @"Data/Programming-Documents/Hyperlink/").LocalPath;
            ImagesDir = new Uri(new Uri(MainDataDir), @"Data/Programming-Documents/Images/").LocalPath;
            JoiningAppendingDir = new Uri(new Uri(MainDataDir), @"Data/Programming-Documents/Joining-Appending/").LocalPath;
            NodeDir = new Uri(new Uri(MainDataDir), @"Data/Programming-Documents/Node/").LocalPath;
            RangeDir = new Uri(new Uri(MainDataDir), @"Data/Programming-Documents/Ranges/").LocalPath;
            SectionsDir = new Uri(new Uri(MainDataDir), @"Data/Programming-Documents/Sections/").LocalPath;
            ShapesDir = new Uri(new Uri(MainDataDir), @"Data/Programming-Documents/Shapes/").LocalPath;
            SignatureDir = new Uri(new Uri(MainDataDir), @"Data/Programming-Documents/Signature/").LocalPath;
            SdtDir = new Uri(new Uri(MainDataDir), @"Data/Programming-Documents/StructuredDocumentTag/").LocalPath;
            StyleDir = new Uri(new Uri(MainDataDir), @"Data/Programming-Documents/Styles/").LocalPath;
            ThemeDir = new Uri(new Uri(MainDataDir), @"Data/Programming-Documents/Theme/").LocalPath;
            WebExtensionsDir = new Uri(new Uri(MainDataDir), @"Data/Programming-Documents/WebExtensions/").LocalPath;
            LoadingSavingDir = new Uri(new Uri(MainDataDir), @"Data/Loading-and-Saving/").LocalPath;
            DatabaseDir = new Uri(new Uri(MainDataDir), @"Data/Database/").LocalPath;
            TablesDir = new Uri(new Uri(MainDataDir), @"Data/Programming-Documents/Tables/").LocalPath;
            LinqDir = new Uri(new Uri(MainDataDir), @"Data/LINQ/").LocalPath;
            MailMergeDir = new Uri(new Uri(MainDataDir), @"Data/Mail-Merge/").LocalPath;
            ViewersVisualizersDir = new Uri(new Uri(MainDataDir), @"Data/Viewers-Visualizers/").LocalPath;
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
        /// Gets the path to the license used by the code examples.
        /// </summary>
        internal static string ViewersVisualizersDir { get; }

        /// <summary>
        /// Gets the path to the documents used by the code examples. Ends with a back slash.
        /// </summary>
        internal static string ArtifactsDir { get; }
    }
}