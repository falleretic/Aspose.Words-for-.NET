using Aspose.Words.Replacing;
using System.Text;
using System.Text.RegularExpressions;
using NUnit.Framework;

namespace Aspose.Words.Examples.CSharp.Programming_Documents.Find_Replace
{
    class ReplaceInHeaderAndFooter : TestDataHelper
    {
        [Test]
        public static void ReplaceTextInFooter()
        {
            //ExStart:ReplaceTextInFooter
            // Open the template document, containing obsolete copyright information in the footer
            Document doc = new Document(DocumentDir + "HeaderFooter.ReplaceText.doc");

            HeaderFooterCollection headersFooters = doc.FirstSection.HeadersFooters;
            HeaderFooter footer = headersFooters[HeaderFooterType.FooterPrimary];

            FindReplaceOptions options = new FindReplaceOptions
            {
                MatchCase = false,
                FindWholeWordsOnly = false
            };

            footer.Range.Replace("(C) 2006 Aspose Pty Ltd.", "Copyright (C) 2020 by Aspose Pty Ltd.", options);

            doc.Save(ArtifactsDir + "HeaderFooter.ReplaceText.doc");
            //ExEnd:ReplaceTextInFooter
        }

        [Test]
        //ExStart:ShowChangesForHeaderAndFooterOrders
        public static void ShowChangesForHeaderAndFooterOrders()
        {
            Document doc = new Document(DocumentDir + "HeaderFooter.ReplaceText.doc");

            // Assert that we use special header and footer for the first page
            // The order for this: first header\footer, even header\footer, primary header\footer
            Section firstPageSection = doc.FirstSection;

            ReplaceLog logger = new ReplaceLog();
            FindReplaceOptions options = new FindReplaceOptions { ReplacingCallback = logger };

            doc.Range.Replace(new Regex("(header|footer)"), "", options);

            doc.Save(ArtifactsDir + "HeaderFooter.HeaderFooterOrder.docx");

            // Prepare our string builder for assert results without "DifferentFirstPageHeaderFooter"
            logger.ClearText();

            // Remove special first page
            // The order for this: primary header, default header, primary footer, default footer, even header\footer
            firstPageSection.PageSetup.DifferentFirstPageHeaderFooter = false;

            doc.Range.Replace(new Regex("(header|footer)"), "", options);
        }

        private class ReplaceLog : IReplacingCallback
        {
            public ReplaceAction Replacing(ReplacingArgs args)
            {
                mTextBuilder.AppendLine(args.MatchNode.GetText());
                return ReplaceAction.Skip;
            }

            internal void ClearText()
            {
                mTextBuilder.Clear();
            }

            internal string Text => mTextBuilder.ToString();

            private readonly StringBuilder mTextBuilder = new StringBuilder();
        }
        // ExEnd:ShowChangesForHeaderAndFooterOrders
    }
}