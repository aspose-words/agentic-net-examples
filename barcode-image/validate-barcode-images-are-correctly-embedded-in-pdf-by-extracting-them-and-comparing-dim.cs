using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Fields;
using Aspose.Words.Drawing;
using Aspose.Words.Replacing;
using Aspose.BarCode.Generation;
using Aspose.Drawing;

public sealed class CustomBarcodeGenerator : IBarcodeGenerator
{
    public Stream GetBarcodeImage(Aspose.Words.Fields.BarcodeParameters parameters)
    {
        MemoryStream stream = new MemoryStream();
        using (BarcodeGenerator generator = new BarcodeGenerator(EncodeTypes.Code128, parameters.BarcodeValue))
        {
            // Set colors if needed
            generator.Parameters.Barcode.BarColor = Color.Black;
            generator.Parameters.BackColor = Color.White;
            generator.Save(stream, BarCodeImageFormat.Png);
        }
        stream.Position = 0;
        return stream;
    }

    public Stream GetOldBarcodeImage(Aspose.Words.Fields.BarcodeParameters parameters)
    {
        return GetBarcodeImage(parameters);
    }
}

public class Program
{
    public static void Main()
    {
        // Create a new document and insert a barcode field.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        FieldDisplayBarcode field = (FieldDisplayBarcode)builder.InsertField(FieldType.FieldDisplayBarcode, true);
        field.BarcodeType = "CODE128";
        field.BarcodeValue = "1234567890";
        field.ForegroundColor = "0x000000";
        field.BackgroundColor = "0xFFFFFF";
        field.ScalingFactor = "100";

        // Register the custom barcode generator for rendered output (PDF).
        doc.FieldOptions.BarcodeGenerator = new CustomBarcodeGenerator();

        // Render the barcode.
        doc.UpdateFields();

        // Save the document as PDF.
        const string pdfPath = "output.pdf";
        doc.Save(pdfPath);

        // Load the generated PDF to extract barcode images.
        Document pdfDoc = new Document(pdfPath);
        Shape barcodeShape = null;
        foreach (Shape shape in pdfDoc.GetChildNodes(NodeType.Shape, true))
        {
            if (shape.HasImage)
            {
                barcodeShape = shape;
                break;
            }
        }

        // Validate that a barcode image was found and has non‑zero dimensions.
        bool isValid = barcodeShape != null && barcodeShape.Width > 0 && barcodeShape.Height > 0;

        // Output validation result.
        Console.WriteLine(isValid
            ? $"Validation passed: Barcode image dimensions = {barcodeShape.Width} x {barcodeShape.Height} points."
            : "Validation failed: No barcode image found or dimensions are invalid.");
    }
}
