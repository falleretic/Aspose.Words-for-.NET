﻿using Aspose.Words.Replacing;
using System.Text.RegularExpressions;

namespace Aspose.Words.Examples.CSharp.Programming_Documents.Find_Replace
{
    class ReplaceWithHTML : TestDataHelper
    {
        public static void Run()
        {
            ReplaceWithHtml();
        }

        // ExStart:ReplaceWithHtml
        public static void ReplaceWithHtml()
        {
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            builder.Writeln("Hello <CustomerName>,");

            FindReplaceOptions options = new FindReplaceOptions();
            options.ReplacingCallback = new ReplaceWithHtmlEvaluator(options);

            doc.Range.Replace(new Regex(@" <CustomerName>,"), string.Empty, options);

            doc.Save(ArtifactsDir + "Range.ReplaceWithInsertHtml.doc");
        }

        private class ReplaceWithHtmlEvaluator : IReplacingCallback
        {
            internal ReplaceWithHtmlEvaluator(FindReplaceOptions options)
            {
                mOptions = options;
            }

            /// <summary>
            /// NOTE: This is a simplistic method that will only work well when the match
            /// starts at the beginning of a run.
            /// </summary>
            ReplaceAction IReplacingCallback.Replacing(ReplacingArgs args)
            {
                DocumentBuilder builder = new DocumentBuilder((Document) args.MatchNode.Document);
                builder.MoveTo(args.MatchNode);

                // Replace '<CustomerName>' text with a red bold name.
                builder.InsertHtml("<b><font color='red'>James Bond, </font></b>");
                args.Replacement = "";

                return ReplaceAction.Replace;
            }

            private readonly FindReplaceOptions mOptions;
        }
        //ExEnd:ReplaceWithHtml
    }
}