using System;
using System.Collections;
using System.Drawing;
using System.Text;
using System.Text.RegularExpressions;
using Aspose.Words.Drawing;
using Aspose.Words.Fields;
using Aspose.Words.Replacing;
using NUnit.Framework;

namespace Aspose.Words.Examples.CSharp
{
    class FindAndReplace : TestDataHelper
    {
        [Test]
        public static void SimpleFindReplace()
        {
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            builder.Writeln("Hello _CustomerName_,");

            // Check the text of the document
            Console.WriteLine("Original document text: " + doc.Range.Text);

            // Replace the text in the document
            doc.Range.Replace("_CustomerName_", "James Bond", new FindReplaceOptions(FindReplaceDirection.Forward));

            // Check the replacement was made
            Console.WriteLine("Document text after replace: " + doc.Range.Text);

            // Save the modified document
            doc.Save(ArtifactsDir + "FindAndReplace.Replace.docx");
        }

        [Test]
        public static void FindAndHighlight()
        {
            //ExStart:FindAndHighlight
            Document doc = new Document(FindReplaceDir + "Find and highlight.docx");

            FindReplaceOptions options = new FindReplaceOptions();
            options.ReplacingCallback = new ReplaceEvaluatorFindAndHighlight();
            options.Direction = FindReplaceDirection.Backward;

            // We want the "your document" phrase to be highlighted
            Regex regex = new Regex("your document", RegexOptions.IgnoreCase);
            doc.Range.Replace(regex, "", options);

            doc.Save(ArtifactsDir + "FindReplaceOptions.FindAndHighlight.docx");
            //ExEnd:FindAndHighlight
        }

        //ExStart:ReplaceEvaluatorFindAndHighlight
        private class ReplaceEvaluatorFindAndHighlight : IReplacingCallback
        {
            /// <summary>
            /// This method is called by the Aspose.Words find and replace engine for each match.
            /// This method highlights the match string, even if it spans multiple runs.
            /// </summary>
            ReplaceAction IReplacingCallback.Replacing(ReplacingArgs e)
            {
                // This is a Run node that contains either the beginning or the complete match
                Node currentNode = e.MatchNode;

                // The first (and may be the only) run can contain text before the match, 
                // in this case it is necessary to split the run
                if (e.MatchOffset > 0)
                    currentNode = SplitRun((Run) currentNode, e.MatchOffset);

                // This array is used to store all nodes of the match for further highlighting
                ArrayList runs = new ArrayList();

                // Find all runs that contain parts of the match string
                int remainingLength = e.Match.Value.Length;
                while (
                    remainingLength > 0 &&
                    currentNode != null &&
                    currentNode.GetText().Length <= remainingLength)
                {
                    runs.Add(currentNode);
                    remainingLength -= currentNode.GetText().Length;

                    // Select the next Run node
                    // Have to loop because there could be other nodes such as BookmarkStart etc.
                    do
                    {
                        currentNode = currentNode.NextSibling;
                    } while (currentNode != null && currentNode.NodeType != NodeType.Run);
                }

                // Split the last run that contains the match if there is any text left
                if (currentNode != null && remainingLength > 0)
                {
                    SplitRun((Run) currentNode, remainingLength);
                    runs.Add(currentNode);
                }

                // Now highlight all runs in the sequence
                foreach (Run run in runs)
                    run.Font.HighlightColor = Color.Yellow;

                // Signal to the replace engine to do nothing because we have already done all what we wanted
                return ReplaceAction.Skip;
            }
        }
        //ExEnd:ReplaceEvaluatorFindAndHighlight

        //ExStart:SplitRun
        /// <summary>
        /// Splits text of the specified run into two runs.
        /// Inserts the new run just after the specified run.
        /// </summary>
        private static Run SplitRun(Run run, int position)
        {
            Run afterRun = (Run) run.Clone(true);
            afterRun.Text = run.Text.Substring(position);
            run.Text = run.Text.Substring(0, position);
            run.ParentNode.InsertAfter(afterRun, run);
            return afterRun;
        }
        //ExEnd:SplitRun

        [Test]
        public static void MetaCharactersInSearchPattern()
        {
            /* meta-characters
            &p - paragraph break
            &b - section break
            &m - page break
            &l - manual line break
            */

            //ExStart:MetaCharactersInSearchPattern
            Document doc = new Document();

            // Use a document builder to add content to the document
            DocumentBuilder builder = new DocumentBuilder(doc);
            builder.Writeln("This is Line 1");
            builder.Writeln("This is Line 2");

            var findReplaceOptions = new FindReplaceOptions();

            doc.Range.Replace("This is Line 1&pThis is Line 2", "This is replaced line", findReplaceOptions);

            builder.MoveToDocumentEnd();
            builder.Write("This is Line 1");
            builder.InsertBreak(BreakType.PageBreak);
            builder.Writeln("This is Line 2");

            doc.Range.Replace("This is Line 1&mThis is Line 2", "Page break is replaced with new text.",
                findReplaceOptions);

            doc.Save(ArtifactsDir + "MetaCharactersInSearchPattern.docx");
            //ExEnd:MetaCharactersInSearchPattern
        }

        [Test]
        public static void ReplaceTextContainingMetaCharacters()
        {
            //ExStart:ReplaceTextContaingMetaCharacters
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            builder.Font.Name = "Arial";
            builder.Writeln("First section");
            builder.Writeln("  1st paragraph");
            builder.Writeln("  2nd paragraph");
            builder.Writeln("{insert-section}");
            builder.Writeln("Second section");
            builder.Writeln("  1st paragraph");

            FindReplaceOptions options = new FindReplaceOptions();
            options.ApplyParagraphFormat.Alignment = ParagraphAlignment.Center;

            // Double each paragraph break after word "section", add kind of underline and make it centered.
            int count = doc.Range.Replace("section&p", "section&p----------------------&p", options);

            // Insert section break instead of custom text tag.
            count = doc.Range.Replace("{insert-section}", "&b", options);

            doc.Save(ArtifactsDir + "ReplaceTextContainingMetaCharacters.docx");
            //ExEnd:ReplaceTextContaingMetaCharacters
        }

        [Test]
        public static void IgnoreTextInsideFields()
        {
            // ExStart:IgnoreTextInsideFields
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Insert field with text inside.
            builder.InsertField("INCLUDETEXT", "Text in field");

            Regex regex = new Regex("e");
            FindReplaceOptions options = new FindReplaceOptions();

            // Replace 'e' in document ignoring text inside field.
            options.IgnoreFields = true;
            doc.Range.Replace(regex, "*", options);
            Console.WriteLine(doc.GetText()); // The output is: \u0013INCLUDETEXT\u0014Text in field\u0015\f

            // Replace 'e' in document NOT ignoring text inside field.
            options.IgnoreFields = false;
            doc.Range.Replace(regex, "*", options);
            Console.WriteLine(doc.GetText()); // The output is: \u0013INCLUDETEXT\u0014T*xt in fi*ld\u0015\f
            // ExEnd:IgnoreTextInsideFields
        }

        [Test]
        public static void IgnoreTextInsideDeleteRevisions()
        {
            // ExStart:IgnoreTextInsideDeleteRevisions
            // Create new document.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Insert non-revised text.
            builder.Writeln("Deleted");
            builder.Write("Text");

            // Remove first paragraph with tracking revisions.
            doc.StartTrackRevisions("author", DateTime.Now);
            doc.FirstSection.Body.FirstParagraph.Remove();
            doc.StopTrackRevisions();

            Regex regex = new Regex("e");
            FindReplaceOptions options = new FindReplaceOptions();

            // Replace 'e' in document ignoring deleted text.
            options.IgnoreDeleted = true;
            doc.Range.Replace(regex, "*", options);
            Console.WriteLine(doc.GetText()); // The output is: Deleted\rT*xt\f

            // Replace 'e' in document NOT ignoring deleted text.
            options.IgnoreDeleted = false;
            doc.Range.Replace(regex, "*", options);
            Console.WriteLine(doc.GetText()); // The output is: D*l*t*d\rT*xt\f
            // ExEnd:IgnoreTextInsideDeleteRevisions
        }

        [Test]
        public static void IgnoreTextInsideInsertRevisions()
        {
            // ExStart:IgnoreTextInsideInsertRevisions
            // Create new document.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Insert text with tracking revisions.
            doc.StartTrackRevisions("author", DateTime.Now);
            builder.Writeln("Inserted");
            doc.StopTrackRevisions();

            // Insert non-revised text.
            builder.Write("Text");

            Regex regex = new Regex("e");
            FindReplaceOptions options = new FindReplaceOptions();

            // Replace 'e' in document ignoring inserted text.
            options.IgnoreInserted = true;
            doc.Range.Replace(regex, "*", options);
            Console.WriteLine(doc.GetText()); // The output is: Inserted\rT*xt\f

            // Replace 'e' in document NOT ignoring inserted text.
            options.IgnoreInserted = false;
            doc.Range.Replace(regex, "*", options);
            Console.WriteLine(doc.GetText()); // The output is: Ins*rt*d\rT*xt\f
            // ExEnd:IgnoreTextInsideInsertRevisions
        }

        [Test]
        public static void ReplaceHtmlTextWithMetaCharacters()
        {
            //ExStart:ReplaceHtmlTextWithMetaCharacters
            Document doc = new Document();

            // Use a document builder to add content to the document
            DocumentBuilder builder = new DocumentBuilder(doc);
            builder.Write("{PLACEHOLDER}");

            FindReplaceOptions findReplaceOptions = new FindReplaceOptions
            {
                ReplacingCallback = new FindAndInsertHtml(),
            };

            doc.Range.Replace("{PLACEHOLDER}", "<p>&ldquo;Some Text&rdquo;</p>", findReplaceOptions);

            doc.Save(ArtifactsDir + "ReplaceHtmlTextWithMetaCharacters.doc");
            //ExEnd:ReplaceHtmlTextWithMetaCharacters
        }

        //ExStart:ReplaceHtmlFindAndInsertHtml
        public sealed class FindAndInsertHtml : IReplacingCallback
        {
            ReplaceAction IReplacingCallback.Replacing(ReplacingArgs e)
            {
                // This is a Run node that contains either the beginning or the complete match
                Node currentNode = e.MatchNode;

                // Create document builder and insert MergeField
                DocumentBuilder builder = new DocumentBuilder(e.MatchNode.Document as Document);
                builder.MoveTo(currentNode);
                builder.InsertHtml(e.Replacement);

                currentNode.Remove();

                // Signal to the replace engine to do nothing because we have already done all what we wanted
                return ReplaceAction.Skip;
            }
        }
        //ExEnd:ReplaceHtmlFindAndInsertHtml

        [Test]
        public static void ReplaceTextInFooter()
        {
            //ExStart:ReplaceTextInFooter
            // Open the template document, containing obsolete copyright information in the footer
            Document doc = new Document(DocumentDir + "Footer.docx");

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
            Document doc = new Document(DocumentDir + "Footer.docx");

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

            private readonly StringBuilder mTextBuilder = new StringBuilder();
        }
        // ExEnd:ShowChangesForHeaderAndFooterOrders

        [Test]
        public static void RReplaceTextWithField()
        {
            Document doc = new Document(FindReplaceDir + "Replace text with fields.docx");

            FindReplaceOptions options = new FindReplaceOptions();
            options.ReplacingCallback = new ReplaceTextWithFieldHandler(FieldType.FieldMergeField);

            // Replace any "PlaceHolderX" instances in the document (where X is a number) with a merge field
            doc.Range.Replace(new Regex(@"PlaceHolder(\d+)"), "", options);

            doc.Save(ArtifactsDir + "Field.ReplaceTextWithFields.doc");
        }


        public class ReplaceTextWithFieldHandler : IReplacingCallback
        {
            public ReplaceTextWithFieldHandler(FieldType type)
            {
                mFieldType = type;
            }

            public ReplaceAction Replacing(ReplacingArgs args)
            {
                ArrayList runs = FindAndSplitMatchRuns(args);

                // Create DocumentBuilder which is used to insert the field
                DocumentBuilder builder = new DocumentBuilder((Document) args.MatchNode.Document);
                builder.MoveTo((Run) runs[runs.Count - 1]);

                // Calculate the name of the field from the FieldType enumeration by removing the first instance of "Field" from the text
                // This works for almost all of the field types
                string fieldName = mFieldType.ToString().ToUpper().Substring(5);

                // Insert the field into the document using the specified field type and the match text as the field name
                // If the fields you are inserting do not require this extra parameter then it can be removed from the string below
                builder.InsertField(string.Format("{0} {1}", fieldName, args.Match.Groups[0]));

                // Now remove all runs in the sequence
                foreach (Run run in runs)
                    run.Remove();

                // Signal to the replace engine to do nothing because we have already done all what we wanted
                return ReplaceAction.Skip;
            }

            /// <summary>
            /// Finds and splits the match runs and returns them in an ArrayList.
            /// </summary>
            public ArrayList FindAndSplitMatchRuns(ReplacingArgs args)
            {
                // This is a Run node that contains either the beginning or the complete match
                Node currentNode = args.MatchNode;

                // The first (and may be the only) run can contain text before the match, 
                // In this case it is necessary to split the run
                if (args.MatchOffset > 0)
                    currentNode = SplitRun((Run) currentNode, args.MatchOffset);

                // This array is used to store all nodes of the match for further removing
                ArrayList runs = new ArrayList();

                // Find all runs that contain parts of the match string
                int remainingLength = args.Match.Value.Length;
                while (
                    remainingLength > 0 &&
                    currentNode != null &&
                    currentNode.GetText().Length <= remainingLength)
                {
                    runs.Add(currentNode);
                    remainingLength -= currentNode.GetText().Length;

                    // Select the next Run node
                    // Have to loop because there could be other nodes such as BookmarkStart etc.
                    do
                    {
                        currentNode = currentNode.NextSibling;
                    } while (currentNode != null && currentNode.NodeType != NodeType.Run);
                }

                // Split the last run that contains the match if there is any text left
                if (currentNode != null && remainingLength > 0)
                {
                    SplitRun((Run) currentNode, remainingLength);
                    runs.Add(currentNode);
                }

                return runs;
            }

            /// <summary>
            /// Splits text of the specified run into two runs.
            /// Inserts the new run just after the specified run.
            /// </summary>
            private static Run SplitRun(Run run, int position)
            {
                Run afterRun = (Run) run.Clone(true);
                afterRun.Text = run.Text.Substring(position);
                run.Text = run.Text.Substring(0, position);
                run.ParentNode.InsertAfter(afterRun, run);
                return afterRun;
            }

            private FieldType mFieldType;
        }

        [Test]
        public static void ReplaceWithEvaluator()
        {
            //ExStart:ReplaceWithEvaluator
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);
            
            builder.Writeln("sad mad bad");

            FindReplaceOptions options = new FindReplaceOptions();
            options.ReplacingCallback = new MyReplaceEvaluator();

            doc.Range.Replace(new Regex("[s|m]ad"), "", options);

            doc.Save(ArtifactsDir + "Range.ReplaceWithEvaluator.doc");
            //ExEnd:ReplaceWithEvaluator
        }

        //ExStart:MyReplaceEvaluator
        private class MyReplaceEvaluator : IReplacingCallback
        {
            /// <summary>
            /// This is called during a replace operation each time a match is found.
            /// This method appends a number to the match string and returns it as a replacement string.
            /// </summary>
            ReplaceAction IReplacingCallback.Replacing(ReplacingArgs e)
            {
                e.Replacement = e.Match.ToString() + mMatchNumber.ToString();
                mMatchNumber++;
                return ReplaceAction.Replace;
            }

            private int mMatchNumber;
        }
        //ExEnd:MyReplaceEvaluator

        [Test]
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

        [Test]
        public static void ReplaceWithRegex()
        {
            //ExStart:ReplaceWithRegex
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);
            
            builder.Writeln("sad mad bad");

            FindReplaceOptions options = new FindReplaceOptions();

            doc.Range.Replace(new Regex("[s|m]ad"), "bad", options);

            doc.Save(ArtifactsDir + "ReplaceWithRegex.doc");
            //ExEnd:ReplaceWithRegex
        }
        
        [Test]
        public static void RecognizeAndSubstitutionsWithinReplacementPatterns()
        {
            // ExStart:RecognizeAndSubstitutionsWithinReplacementPatterns
            // Create new document.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Write some text.
            builder.Write("Jason give money to Paul.");

            Regex regex = new Regex(@"([A-z]+) give money to ([A-z]+)");

            // Replace text using substitutions.
            FindReplaceOptions options = new FindReplaceOptions();
            options.UseSubstitutions = true;
            doc.Range.Replace(regex, @"$2 take money from $1", options);
            // ExEnd:RecognizeAndSubstitutionsWithinReplacementPatterns
        }

        [Test]
        public static void ReplaceWithString()
        {
            //ExStart:ReplaceWithString
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);
            
            builder.Writeln("sad mad bad");

            doc.Range.Replace("sad", "bad", new FindReplaceOptions(FindReplaceDirection.Forward));

            doc.Save(ArtifactsDir + "ReplaceWithString.doc");
            //ExEnd:ReplaceWithString
        }

        [Test]
        //ExStart:FineReplaceUsingLegacyOrder
        public static void FineReplaceUsingLegacyOrder()
        {
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Insert 3 tags to appear in sequential order, the second of which will be inside a text box
            builder.Writeln("[tag 1]");
            Shape textBox = builder.InsertShape(ShapeType.TextBox, 100, 50);
            builder.Writeln("[tag 3]");

            builder.MoveTo(textBox.FirstParagraph);
            builder.Write("[tag 2]");

            FindReplaceOptions options = new FindReplaceOptions();
            options.ReplacingCallback = new UsingLegacyOrder.ReplacingCallback();
            options.UseLegacyOrder = true;

            doc.Range.Replace(new Regex(@"\[(.*?)\]"), "", options);

            doc.Save(ArtifactsDir + "FineReplaceUsingLegacyOrder.docx");
        }

        private class ReplacingCallback : IReplacingCallback
        {
            ReplaceAction IReplacingCallback.Replacing(ReplacingArgs e)
            {
                Console.Write(e.Match.Value);
                return ReplaceAction.Replace;
            }
        }
        //ExEnd:FineReplaceUsingLegacyOrder
    }
}
