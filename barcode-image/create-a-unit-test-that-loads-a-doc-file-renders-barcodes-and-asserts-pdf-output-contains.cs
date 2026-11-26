using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Fields;
using Aspose.Words.Drawing;
using Aspose.BarCode.Generation;
using Aspose.Drawing;

public sealed class CustomBarcodeGenerator : IBarcodeGenerator
{
    public Stream GetBarcodeImage(Aspose.Words.Fields.BarcodeParameters parameters)
    {
        var stream = new MemoryStream();
        using (var generator = new BarcodeGenerator(EncodeTypes.Code128, parameters.BarcodeValue))
        {
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
        var doc = new Document();
        var builder = new DocumentBuilder(doc);
        var field = (FieldDisplayBarcode)builder.InsertField(FieldType.FieldDisplayBarcode, true);
        field.BarcodeType = "CODE128";
        field.BarcodeValue = "1234567890";
        field.ForegroundColor = "0x000000";
        field.BackgroundColor = "0xFFFFFF";

        // Register custom barcode generator for rendering.
        doc.FieldOptions.BarcodeGenerator = new CustomBarcodeGenerator();

        // Update fields to generate barcode image.
        doc.UpdateFields();

        // Save as PDF.
        const string pdfPath = "output.pdf";
        doc.Save(pdfPath);

        // Load the PDF and verify that it contains at least one image shape.
        var pdfDoc = new Document(pdfPath);
        int imageCount = 0;
        foreach (Node node in pdfDoc.GetChildNodes(NodeType.Shape, true))
        {
            var shape = (Shape)node;
            if (shape.HasImage)
                imageCount++;
        }

        if (imageCount > 0)
        {
            Console.WriteLine("Test Passed: PDF contains barcode image.");
        }
        else
        {
            Console.WriteLine("Test Failed: No barcode image found in PDF.");
        }
    }
}
