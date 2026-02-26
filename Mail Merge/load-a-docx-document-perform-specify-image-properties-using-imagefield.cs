using System;
using System.Data;
using Aspose.Words;
using Aspose.Words.MailMerging;
using Aspose.Words.Saving;
using Aspose.Words.Fields; // Needed for MergeFieldImageDimension and its enum

class Program
{
    static void Main()
    {
        // Load the existing DOCX document.
        Document doc = new Document("Input.docx");

        // Attach a callback that will set image dimensions during mail merge.
        doc.MailMerge.FieldMergingCallback = new ImageSizeCallback(
            width: 150,               // Desired width.
            height: 150,              // Desired height.
            unit: MergeFieldImageDimensionUnit.Point);

        // Prepare a simple data source containing image file names.
        DataTable data = new DataTable();
        data.Columns.Add("Photo");               // Column name must match the merge field (e.g., MERGEFIELD Photo).
        data.Rows.Add("Image1.jpg");
        data.Rows.Add("Image2.png");

        // Perform the mail merge; the callback will adjust each image.
        doc.MailMerge.Execute(data);

        // Save the resulting document as PDF.
        PdfSaveOptions pdfOptions = new PdfSaveOptions(); // Default options; customize if needed.
        doc.Save("Output.pdf", pdfOptions);
    }
}

// Callback that specifies image file name and dimensions for each image merge field.
class ImageSizeCallback : IFieldMergingCallback
{
    private readonly double _width;
    private readonly double _height;
    private readonly MergeFieldImageDimensionUnit _unit;

    public ImageSizeCallback(double width, double height, MergeFieldImageDimensionUnit unit)
    {
        _width = width;
        _height = height;
        _unit = unit;
    }

    // Not used for non‑image fields.
    public void FieldMerging(FieldMergingArgs e) { }

    // Called for each image merge field.
    public void ImageFieldMerging(ImageFieldMergingArgs e)
    {
        // Use the field value (expected to be a file path) as the image source.
        e.ImageFileName = e.FieldValue?.ToString();

        // Set the desired dimensions using MergeFieldImageDimension.
        e.ImageWidth = new MergeFieldImageDimension(_width, _unit);
        e.ImageHeight = new MergeFieldImageDimension(_height, _unit);
    }
}
