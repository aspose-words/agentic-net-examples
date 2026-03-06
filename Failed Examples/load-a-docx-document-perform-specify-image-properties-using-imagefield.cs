// ALL ATTEMPTS FAILED. Below is the last generated code.

using System;
using System.Data;
using Aspose.Words;
using Aspose.Words.MailMerging;
using Aspose.Words.Saving;

namespace AsposeWordsImageMergeExample
{
    // Callback that customizes image merging behavior.
    // Implements IFieldMergingCallback to handle image merge fields.
    public class ImageMergeCallback : IFieldMergingCallback
    {
        private readonly double _imageWidth;
        private readonly double _imageHeight;
        private readonly MergeFieldImageDimensionUnit _unit;

        public ImageMergeCallback(double imageWidth, double imageHeight, MergeFieldImageDimensionUnit unit)
        {
            _imageWidth = imageWidth;
            _imageHeight = imageHeight;
            _unit = unit;
        }

        // Not used for non‑image fields.
        public void FieldMerging(FieldMergingArgs args)
        {
            // No custom processing required.
        }

        // Called for each image merge field.
        public void ImageFieldMerging(ImageFieldMergingArgs args)
        {
            // Use the field value (expected to be a file name) as the source image.
            args.ImageFileName = args.FieldValue?.ToString();

            // Set desired dimensions for the inserted image.
            args.ImageWidth = new MergeFieldImageDimension(_imageWidth, _unit);
            args.ImageHeight = new MergeFieldImageDimension(_imageHeight, _unit);

            // If you need to insert a pre‑configured Shape instead of an image file,
            // you could assign args.Shape here and skip the other properties.
        }
    }

    class Program
    {
        static void Main()
        {
            // Path to the source DOCX containing an image MERGEFIELD (e.g., MERGEFIELD Image:Photo).
            const string inputPath = "input.docx";

            // Load the document using the Document(string) constructor.
            Document doc = new Document(inputPath);

            // Prepare a data source with image file names.
            DataTable table = new DataTable("Images");
            table.Columns.Add(new DataColumn("Photo"));
            table.Rows.Add(@"Images\Sample1.jpg");
            table.Rows.Add(@"Images\Sample2.png");
            table.Rows.Add(@"Images\Sample3.emf");

            // Attach the callback that will set image dimensions during the merge.
            doc.MailMerge.FieldMergingCallback = new ImageMergeCallback(
                imageWidth: 200,          // Desired width
                imageHeight: 200,         // Desired height
                unit: MergeFieldImageDimensionUnit.Point);

            // Execute the mail merge. The field name in the document must be "Image:Photo".
            doc.MailMerge.Execute(table);

            // Optional: update any remaining fields (e.g., TOC) before saving.
            doc.UpdateFields();

            // Configure PDF save options (e.g., keep default color mode).
            PdfSaveOptions pdfOptions = new PdfSaveOptions
            {
                // Example: set the PDF to use high‑quality rendering.
                UseHighQualityRendering = true
            };

            // Save the merged document as PDF using the Save(string, SaveOptions) overload.
            const string outputPath = "output.pdf";
            doc.Save(outputPath, pdfOptions);
        }
    }
}
