using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Fields;
using Aspose.Words.Saving;
using Aspose.Words.Drawing;
using Aspose.BarCode.Generation;

public sealed class CustomBarcodeGenerator : IBarcodeGenerator
{
    // Generates a PNG barcode image based on the provided parameters.
    public Stream GetBarcodeImage(Aspose.Words.Fields.BarcodeParameters parameters)
    {
        var stream = new MemoryStream();
        using (var generator = new BarcodeGenerator(EncodeTypes.Code128, parameters.BarcodeValue))
        {
            // Optional: set colors if needed
            // generator.Parameters.Barcode.BarColor = Aspose.Drawing.Color.Black;
            // generator.Parameters.BackColor = Aspose.Drawing.Color.White;
            generator.Save(stream, BarCodeImageFormat.Png);
        }
        stream.Position = 0;
        return stream;
    }

    // Legacy method – delegates to the main implementation.
    public Stream GetOldBarcodeImage(Aspose.Words.Fields.BarcodeParameters parameters)
    {
        return GetBarcodeImage(parameters);
    }
}

public class Program
{
    public static void Main()
    {
        // Temporary file path for the generated PDF.
        string pdfPath = Path.Combine(Path.GetTempPath(), "BarcodeTest.pdf");

        // Create a new document and insert a DISPLAYBARCODE field.
        var doc = new Document();
        var builder = new DocumentBuilder(doc);
        var field = (FieldDisplayBarcode)builder.InsertField(FieldType.FieldDisplayBarcode, true);
        field.BarcodeType = "CODE128";
        field.BarcodeValue = "1234567890";
        field.ForegroundColor = "0x000000";
        field.BackgroundColor = "0xFFFFFF";

        // Register the custom barcode generator for rendered output (PDF).
        doc.FieldOptions.BarcodeGenerator = new CustomBarcodeGenerator();

        // Render the barcode field.
        doc.UpdateFields();

        // Save the document as PDF.
        doc.Save(pdfPath, SaveFormat.Pdf);

        // Load the generated PDF for verification.
        var pdfDoc = new Document(pdfPath);

        // Count image shapes (the barcode should be rendered as an image).
        int imageCount = 0;
        foreach (Node node in pdfDoc.GetChildNodes(NodeType.Shape, true))
        {
            if (node is Shape shape && shape.HasImage)
                imageCount++;
        }

        // Simple assertion: PDF must contain at least one image (the barcode).
        if (imageCount == 0)
            throw new InvalidOperationException("PDF does not contain any rendered barcode images.");

        // Clean up temporary PDF file.
        if (File.Exists(pdfPath))
            File.Delete(pdfPath);
    }
}
