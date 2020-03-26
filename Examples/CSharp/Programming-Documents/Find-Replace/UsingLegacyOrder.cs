using Aspose.Words.Replacing;
using System;
using System.Text.RegularExpressions;

namespace Aspose.Words.Examples.CSharp.Programming_Documents.Find_and_Replace
{
    class UsingLegacyOrder
    {
        public static void Run()
        {
            // The path to the documents directory.
            string dataDir = RunExamples.GetDataDir_FindAndReplace();

            FineReplaceUsingLegacyOrder(dataDir);
        }

        // ExStart:FineReplaceUsingLegacyOrder
        public static void FineReplaceUsingLegacyOrder(string dataDir)
        {
            // Open the document.
            Document doc = new Document(@"source.docx");
            FindReplaceOptions options = new FindReplaceOptions();
            options.ReplacingCallback = new ReplacingCallback();
            options.UseLegacyOrder = true;

            doc.Range.Replace(new Regex(@"\[(.*?)\]"), "", options);

            dataDir = dataDir + "usingLegacyOrder_out.doc";
            doc.Save(dataDir);
        }

        private class ReplacingCallback : IReplacingCallback
        {
            ReplaceAction IReplacingCallback.Replacing(ReplacingArgs e)
            {
                Console.Write(e.Match.Value);
                return ReplaceAction.Replace;
            }
        }

        // ExEnd:FineReplaceUsingLegacyOrder
    }
}