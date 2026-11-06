using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Fields;
using Aspose.Words.Drawing;
using Aspose.Words.Loading;
using Aspose.Words.Saving;
using Aspose.Words.Replacing;
using Aspose.BarCode.Generation;
using Aspose.Drawing;

public sealed class CustomBarcodeGenerator : IBarcodeGenerator
{
    // Generates a barcode image stream based on the provided parameters.
    public Stream GetBarcodeImage(Aspose.Words.Fields.BarcodeParameters parameters)
    {
        var stream = new MemoryStream();
        using (var generator = new BarcodeGenerator(EncodeTypes.Code128, parameters.BarcodeValue))
        {
            generator.Save(stream, BarCodeImageFormat.Png);
        }
        stream.Position = 0;
        return stream;
    }

    // Legacy method – delegate to the main implementation.
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

        // Register the custom barcode generator for rendered output (PDF).
        doc.FieldOptions.BarcodeGenerator = new CustomBarcodeGenerator();

        // Update fields to render the barcode.
        doc.UpdateFields();

        // Save the document as PDF.
        const string pdfPath = "output.pdf";
        doc.Save(pdfPath, SaveFormat.Pdf);

        // Load the generated PDF back into a document.
        Document pdfDoc = new Document(pdfPath);

        // Find the first image shape (the rendered barcode).
        Shape barcodeShape = null;
        NodeCollection shapes = pdfDoc.GetChildNodes(NodeType.Shape, true);
        foreach (Shape shape in shapes)
        {
            if (shape.HasImage)
            {
                barcodeShape = shape;
                break;
            }
        }

        if (barcodeShape == null)
        {
            Console.WriteLine("Barcode image not found in PDF.");
            return;
        }

        // Extract the image bytes from the shape.
        byte[] pdfImageBytes = barcodeShape.ImageData.ImageBytes;

        // Load the image using Aspose.Drawing to get its dimensions.
        float pdfWidth, pdfHeight;
        using (var ms = new MemoryStream(pdfImageBytes))
        using (var img = Image.FromStream(ms))
        {
            pdfWidth = img.Width;
            pdfHeight = img.Height;
        }

        // Generate the original barcode image using the same generator.
        byte[] originalImageBytes;
        using (var genStream = new MemoryStream())
        {
            using (var generator = new BarcodeGenerator(EncodeTypes.Code128, "1234567890"))
            {
                generator.Save(genStream, BarCodeImageFormat.Png);
            }
            originalImageBytes = genStream.ToArray();
        }

        // Load the original image to obtain its dimensions.
        float originalWidth, originalHeight;
        using (var ms = new MemoryStream(originalImageBytes))
        using (var img = Image.FromStream(ms))
        {
            originalWidth = img.Width;
            originalHeight = img.Height;
        }

        // Compute aspect ratios.
        double pdfAspect = pdfWidth / pdfHeight;
        double originalAspect = originalWidth / originalHeight;
        double tolerance = 0.01; // Allow 1% difference.

        if (Math.Abs(pdfAspect - originalAspect) <= tolerance)
        {
            Console.WriteLine("Aspect ratio matches.");
        }
        else
        {
            Console.WriteLine($"Aspect ratio mismatch. PDF: {pdfAspect:F4}, Original: {originalAspect:F4}");
        }
    }
}
