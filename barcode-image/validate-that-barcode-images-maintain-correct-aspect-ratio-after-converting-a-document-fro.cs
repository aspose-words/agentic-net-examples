using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Fields;
using Aspose.Words.Drawing;
using Aspose.Words.Saving;
using Aspose.BarCode.Generation;
using Aspose.Drawing;

public sealed class CustomBarcodeGenerator : IBarcodeGenerator
{
    // Generates a PNG barcode image based on the provided parameters.
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

    // Legacy method – delegates to the main generator.
    public Stream GetOldBarcodeImage(Aspose.Words.Fields.BarcodeParameters parameters)
    {
        return GetBarcodeImage(parameters);
    }
}

public class Program
{
    public static void Main()
    {
        // Create a new document and insert a DISPLAYBARCODE field.
        var doc = new Document();
        var builder = new DocumentBuilder(doc);
        var field = (FieldDisplayBarcode)builder.InsertField(FieldType.FieldDisplayBarcode, true);
        field.BarcodeType = "CODE128";
        field.BarcodeValue = "1234567890";

        // Register the custom barcode generator for rendered output (PDF).
        doc.FieldOptions.BarcodeGenerator = new CustomBarcodeGenerator();

        // Render the barcode.
        doc.UpdateFields();

        // Save as PDF.
        const string pdfPath = "output.pdf";
        doc.Save(pdfPath, SaveFormat.Pdf);

        // Load the PDF to inspect the rendered barcode image.
        var pdfDoc = new Document(pdfPath);
        Shape barcodeShape = null;
        var shapes = pdfDoc.GetChildNodes(NodeType.Shape, true);
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

        // Extract the rendered barcode image from the PDF.
        byte[] pdfImageBytes = barcodeShape.ImageData.ImageBytes;
        double pdfAspectRatio;
        using (var ms = new MemoryStream(pdfImageBytes))
        using (var pdfImage = Image.FromStream(ms))
        {
            pdfAspectRatio = (double)pdfImage.Width / pdfImage.Height;
        }

        // Generate the expected barcode image using the custom generator.
        var parameters = new Aspose.Words.Fields.BarcodeParameters
        {
            BarcodeType = "CODE128",
            BarcodeValue = "1234567890"
        };
        double expectedAspectRatio;
        using (var genStream = new CustomBarcodeGenerator().GetBarcodeImage(parameters))
        using (var genImage = Image.FromStream(genStream))
        {
            expectedAspectRatio = (double)genImage.Width / genImage.Height;
        }

        // Compare aspect ratios with a tolerance.
        const double tolerance = 0.01;
        bool aspectRatioValid = Math.Abs(pdfAspectRatio - expectedAspectRatio) <= tolerance;

        Console.WriteLine($"PDF barcode aspect ratio: {pdfAspectRatio:F4}");
        Console.WriteLine($"Expected barcode aspect ratio: {expectedAspectRatio:F4}");
        Console.WriteLine($"Aspect ratio valid: {aspectRatioValid}");
    }
}
