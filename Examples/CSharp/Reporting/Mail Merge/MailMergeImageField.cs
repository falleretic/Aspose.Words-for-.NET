using Aspose.Words.Drawing;
using Aspose.Words.MailMerging;
using NUnit.Framework;

namespace Aspose.Words.Examples.CSharp
{
    class MailMergeImageField : TestDataHelper
    {
        [Test]
        public static void Run()
        {
            // ExStart:MailMergeImageField       
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            builder.Writeln("{{#foreach example}}");
            builder.Writeln("{{Image(126pt;126pt):stempel}}");
            builder.Writeln("{{/foreach example}}");

            doc.MailMerge.UseNonMergeFields = true;
            doc.MailMerge.TrimWhitespaces = true;
            doc.MailMerge.UseWholeParagraphAsRegion = false;
            doc.MailMerge.CleanupOptions = MailMergeCleanupOptions.RemoveEmptyTableRows
                    | MailMergeCleanupOptions.RemoveContainingFields
                    | MailMergeCleanupOptions.RemoveUnusedRegions
                    | MailMergeCleanupOptions.RemoveUnusedFields;

            // Add a handler for the MergeField event.
            doc.MailMerge.FieldMergingCallback = new ImageFieldMergingHandler();
            doc.MailMerge.ExecuteWithRegions(new DataSourceRoot());

            doc.Save(ArtifactsDir + "MailMerge.ImageMailMerge.docx");
            // ExEnd:MailMergeImageField
        }

        // ExStart:ImageFieldMergingHandler
        private class ImageFieldMergingHandler : IFieldMergingCallback
        {
            void IFieldMergingCallback.FieldMerging(FieldMergingArgs args)
            {
                //  Implementation is not required.
            }

            void IFieldMergingCallback.ImageFieldMerging(ImageFieldMergingArgs args)
            {
                Shape shape = new Shape(args.Document, ShapeType.Image);
                shape.Width = 126;
                shape.Height = 126;
                shape.WrapType = WrapType.Square;

                shape.ImageData.SetImage(MyDir + "Mail merge image.png");

                args.Shape = shape;
            }
        }
        // ExEnd:ImageFieldMergingHandler

        // ExStart:DataSourceRoot
        public class DataSourceRoot : IMailMergeDataSourceRoot
        {
            public IMailMergeDataSource GetDataSource(string s)
            {
                return new DataSource();
            }

            private class DataSource : IMailMergeDataSource
            {
                private bool next = true;

                string IMailMergeDataSource.TableName => TableName();

                private static string TableName()
                {
                    return "example";
                }

                public bool MoveNext()
                {
                    bool result = next;
                    next = false;
                    return result;
                }

                public IMailMergeDataSource GetChildDataSource(string s)
                {
                    return null;
                }

                public bool GetValue(string fieldName, out object fieldValue)
                {
                    fieldValue = null;
                    return false;
                }
            }
        }
        // ExEnd:DataSourceRoot
    }
}
