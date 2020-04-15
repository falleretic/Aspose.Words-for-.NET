using System.Text.RegularExpressions;
using Aspose.Words.Replacing;
using NUnit.Framework;

namespace Aspose.Words.Examples.CSharp.Programming_Documents.Find_and_Replace
{
    class ReplaceWithEvaluator : TestDataHelper
    {
        [Test]
        public static void Run()
        {
            //ExStart:ReplaceWithEvaluator
            Document doc = new Document(FindReplaceDir + "Range.ReplaceWithEvaluator.doc");

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
    }
}