using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Fields;
using Aspose.BarCode.Generation;
using Aspose.Drawing;

public sealed class CustomBarcodeGenerator : IBarcodeGenerator
{
    // Generates a PNG barcode image based on the parameters supplied by the DISPLAYBARCODE field.
    public Stream GetBarcodeImage(Aspose.Words.Fields.BarcodeParameters parameters)
    {
        var stream = new MemoryStream();
        // For simplicity we always use Code128; in real scenarios map parameters.BarcodeType to an EncodeTypes value.
        using (var generator = new BarcodeGenerator(EncodeTypes.Code128, parameters.BarcodeValue))
        {
            generator.Save(stream, BarCodeImageFormat.Png);
        }
        stream.Position = 0;
        return stream;
    }

    // Legacy method – delegate to the primary implementation.
    public Stream GetOldBarcodeImage(Aspose.Words.Fields.BarcodeParameters parameters)
    {
        return GetBarcodeImage(parameters);
    }
}

public class Program
{
    public static void Main()
    {
        // -----------------------------------------------------------------
        // Create a simple DOCX template containing a DISPLAYBARCODE field.
        // -----------------------------------------------------------------
        var templatePath = "template.docx";
        var templateDoc = new Document();
        var builder = new DocumentBuilder(templateDoc);

        // Insert a typed DISPLAYBARCODE field.
        var field = (FieldDisplayBarcode)builder.InsertField(FieldType.FieldDisplayBarcode, true);
        field.BarcodeType = "CODE128";
        field.BarcodeValue = "PLACEHOLDER";
        field.ForegroundColor = "0x000000";
        field.BackgroundColor = "0xFFFFFF";

        templateDoc.Save(templatePath);

        // -----------------------------------------------------------------
        // Load the template, populate the barcode field, and render to PDF.
        // -----------------------------------------------------------------
        var doc = new Document(templatePath);

        // Register the custom barcode generator – required for PDF rendering.
        doc.FieldOptions.BarcodeGenerator = new CustomBarcodeGenerator();

        // Populate DISPLAYBARCODE fields with actual data.
        foreach (Field fieldItem in doc.Range.Fields)
        {
            if (fieldItem.Type == FieldType.FieldDisplayBarcode)
            {
                var barcodeField = (FieldDisplayBarcode)fieldItem;
                barcodeField.BarcodeType = "QR";
                barcodeField.BarcodeValue = "https://example.com";
                barcodeField.ForegroundColor = "0x0000FF";
                barcodeField.BackgroundColor = "0xFFFFFF";
            }
        }

        // Render the barcodes.
        doc.UpdateFields();

        // Save the final document as PDF.
        doc.Save("output.pdf");
    }
}
