﻿using NUnit.Framework;

namespace Aspose.Words.Examples.CSharp.Ranges
{
    class RangesGetText : TestDataHelper
    {
        [Test]
        public static void Run()
        {
            //ExStart:RangesGetText
            Document doc = new Document(RangeDir + "Document.doc");
            string text = doc.Range.Text;
            //ExEnd:RangesGetText
        }
    }
}