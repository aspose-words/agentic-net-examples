using System;
using System.Data;
using Aspose.Words;
using Aspose.Words.MailMerging;
using Aspose.Words.Fields;

class Program
{
    static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert an image merge field that will receive images during mail merge.
        // The field name must start with "Image:" to be recognized as an image field.
        builder.InsertField("MERGEFIELD Image:Photo");

        // Prepare a simple data source containing file paths to the images.
        DataTable table = new DataTable("Images");
        table.Columns.Add("Photo");
        table.Rows.Add(@"C:\Images\Logo.jpg");
        table.Rows.Add(@"C:\Images\Sample.png");

        // Attach a callback that will resize every merged image to the desired dimensions.
        doc.MailMerge.FieldMergingCallback = new ImageResizer(150, 150, MergeFieldImageDimensionUnit.Point);

        // Perform the mail merge.
        doc.MailMerge.Execute(table);

        // Save the resulting document.
        doc.Save("ResizedImages.docx");
    }

    // Implementation of IFieldMergingCallback that resizes images.
    private class ImageResizer : IFieldMergingCallback
    {
        private readonly double _width;
        private readonly double _height;
        private readonly MergeFieldImageDimensionUnit _unit;

        public ImageResizer(double width, double height, MergeFieldImageDimensionUnit unit)
        {
            _width = width;
            _height = height;
            _unit = unit;
        }

        // Not used for plain text fields.
        public void FieldMerging(FieldMergingArgs args) { }

        // Called for each image merge field.
        public void ImageFieldMerging(ImageFieldMergingArgs args)
        {
            // Supply the image file name from the data source.
            args.ImageFileName = args.FieldValue.ToString();

            // Override the default dimensions with the desired size.
            args.ImageWidth = new MergeFieldImageDimension(_width, _unit);
            args.ImageHeight = new MergeFieldImageDimension(_height, _unit);
        }
    }
}
