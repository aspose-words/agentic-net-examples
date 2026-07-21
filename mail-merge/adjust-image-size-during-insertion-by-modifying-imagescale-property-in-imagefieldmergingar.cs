using System;
using System.Data;
using System.IO;
using Aspose.Words;
using Aspose.Words.MailMerging;
using Aspose.Words.Fields;

public class Program
{
    public static void Main()
    {
        // Create a new empty document.
        Document doc = new Document();

        // Insert a MERGEFIELD that will accept an image during mail merge.
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.InsertField("MERGEFIELD Image:Photo");

        // Prepare a simple data source. The actual value is not used because the callback
        // creates an in‑memory image, but the column must exist.
        DataTable table = new DataTable("Images");
        table.Columns.Add(new DataColumn("Photo"));
        table.Rows.Add("ignored1");
        table.Rows.Add("ignored2");

        // Set a callback that will insert a generated image and resize it to 150x150 points.
        doc.MailMerge.FieldMergingCallback = new MergedImageResizer(150, 150, MergeFieldImageDimensionUnit.Point);

        // Execute the mail merge.
        doc.MailMerge.Execute(table);

        // Update fields to reflect the merged content.
        doc.UpdateFields();

        // Save the resulting document.
        doc.Save("MergedImages.docx");
    }

    // Callback that creates an image and resizes it during mail merge.
    private class MergedImageResizer : IFieldMergingCallback
    {
        private readonly double _width;
        private readonly double _height;
        private readonly MergeFieldImageDimensionUnit _unit;

        public MergedImageResizer(double width, double height, MergeFieldImageDimensionUnit unit)
        {
            _width = width;
            _height = height;
            _unit = unit;
        }

        // No custom processing required for text fields.
        public void FieldMerging(FieldMergingArgs args)
        {
            // Intentionally left blank.
        }

        // Called for each image merge field.
        public void ImageFieldMerging(ImageFieldMergingArgs args)
        {
            // A minimal 1x1 pixel PNG (transparent) encoded in base64.
            const string base64Png = "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAwMCAO+XK2cAAAAASUVORK5CYII=";
            byte[] pngBytes = Convert.FromBase64String(base64Png);

            // Provide the image via a stream. Aspose.Words will read the image from this stream.
            args.ImageStream = new MemoryStream(pngBytes);

            // Override the size of the inserted image.
            args.ImageWidth = new MergeFieldImageDimension(_width, _unit);
            args.ImageHeight = new MergeFieldImageDimension(_height, _unit);
        }
    }
}
