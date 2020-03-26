using Aspose.Words.Fields;
using System;

namespace Aspose.Words.Examples.CSharp.Programming_Documents.Working_with_Fields
{
    class EvaluateIFCondition
    {
        public static void Run()
        {
            // ExStart:EvaluateIFCondition
            DocumentBuilder builder = new DocumentBuilder();
            FieldIf field = (FieldIf) builder.InsertField("IF 1 = 1", null);

            FieldIfComparisonResult actualResult = field.EvaluateCondition();
            Console.WriteLine(actualResult);
            // ExEnd:EvaluateIFCondition

            Console.WriteLine("\nEvaluates the IF condition successfully.");
        }
    }
}