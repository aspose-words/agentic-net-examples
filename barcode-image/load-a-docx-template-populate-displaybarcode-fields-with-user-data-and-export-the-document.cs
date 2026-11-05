using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Fields;
using Aspose.BarCode.Generation;

public sealed class CustomBarcodeGenerator : IBarcodeGenerator
{
    // Generates a PNG barcode image based on the parameters supplied by Aspose.Words.
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

    // Legacy method – simply forwards to the main implementation.
    public Stream GetOldBarcodeImage(Aspose.Words.Fields.BarcodeParameters parameters)
    {
        return GetBarcodeImage(parameters);
    }
}

public class Program
{
    public static void Main()
    {
        const string templatePath = "template.docx";

        // Create a minimal template if it does not exist.
        if (!File.Exists(templatePath))
        {
            var tempDoc = new Document();
            var builder = new DocumentBuilder(tempDoc);
            // Insert a DISPLAYBARCODE field placeholder.
            var placeholderField = (FieldDisplayBarcode)builder.InsertField(FieldType.FieldDisplayBarcode, true);
            placeholderField.BarcodeType = "CODE128";
            placeholderField.BarcodeValue = "PLACEHOLDER";
            tempDoc.Save(templatePath);
        }

        // Load the DOCX template.
        var doc = new Document(templatePath);

        // Register the custom barcode generator for rendered output (PDF).
        doc.FieldOptions.BarcodeGenerator = new CustomBarcodeGenerator();

        // Populate each DISPLAYBARCODE field with the desired data.
        foreach (Field field in doc.Range.Fields)
        {
            if (field.Type == FieldType.FieldDisplayBarcode)
            {
                var barcodeField = (FieldDisplayBarcode)field;
                barcodeField.BarcodeType = "CODE128";
                barcodeField.BarcodeValue = "1234567890";
                // Optional color settings (hex string format).
                // barcodeField.ForegroundColor = "0x000000";
                // barcodeField.BackgroundColor = "0xFFFFFF";
            }
        }

        // Render the barcodes.
        doc.UpdateFields();

        // Save the document as PDF.
        doc.Save("output.pdf");
    }
}
