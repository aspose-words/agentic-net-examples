using System;
using System.Data;
using System.IO;
using Aspose.Words;
using Aspose.Words.Fields;
using Aspose.Words.MailMerging;

namespace ImageMailMergeExample
{
    public class Program
    {
        public static void Main()
        {
            // Create a temporary 1x1 PNG image without using System.Drawing.
            string imagePath = Path.Combine(Path.GetTempPath(), "tempImage.png");
            // Minimal PNG byte array (transparent 1x1 pixel).
            byte[] pngBytes = new byte[]
            {
                0x89,0x50,0x4E,0x47,0x0D,0x0A,0x1A,0x0A,
                0x00,0x00,0x00,0x0D,0x49,0x48,0x44,0x52,
                0x00,0x00,0x00,0x01,0x00,0x00,0x00,0x01,
                0x08,0x06,0x00,0x00,0x00,0x1F,0x15,0xC4,
                0x89,0x00,0x00,0x00,0x0A,0x49,0x44,0x41,
                0x54,0x78,0x9C,0x63,0x00,0x01,0x00,0x00,
                0x05,0x00,0x01,0x0D,0x0A,0x2D,0xB4,0x00,
                0x00,0x00,0x00,0x49,0x45,0x4E,0x44,0xAE,
                0x42,0x60,0x82
            };
            File.WriteAllBytes(imagePath, pngBytes);

            // Create a new document and insert an image merge field.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);
            builder.InsertField("MERGEFIELD Image:ImageColumn");

            // Build a data source with the image file name.
            DataTable table = new DataTable("Images");
            table.Columns.Add(new DataColumn("ImageColumn"));
            table.Rows.Add(imagePath);

            // Set up a callback to resize the merged image.
            doc.MailMerge.FieldMergingCallback = new MergedImageResizer(100, 100, MergeFieldImageDimensionUnit.Point);
            doc.MailMerge.Execute(table);
            doc.UpdateFields();

            // Save the result.
            string outputPath = Path.Combine(Environment.CurrentDirectory, "output.docx");
            doc.Save(outputPath);
        }

        // Callback that sets the image dimensions during mail merge.
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

            // No custom text handling required.
            public void FieldMerging(FieldMergingArgs args) { }

            public void ImageFieldMerging(ImageFieldMergingArgs args)
            {
                // Provide the image file name from the data source.
                args.ImageFileName = args.FieldValue.ToString();

                // Override the image size.
                args.ImageWidth = new MergeFieldImageDimension(_width, _unit);
                args.ImageHeight = new MergeFieldImageDimension(_height, _unit);
            }
        }
    }
}
