using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Fields;
using Aspose.BarCode.Generation;
using Aspose.Drawing;

public sealed class CustomBarcodeGenerator : IBarcodeGenerator
{
    // Generates a barcode image using Aspose.BarCode and returns it as a PNG stream.
    public Stream GetBarcodeImage(Aspose.Words.Fields.BarcodeParameters parameters)
    {
        var stream = new MemoryStream();
        using (var generator = new BarcodeGenerator(EncodeTypes.Code128, parameters.BarcodeValue))
        {
            // Set barcode colors.
            generator.Parameters.Barcode.BarColor = Color.Black;
            generator.Parameters.BackColor = Color.White;

            // Save the barcode as PNG to the stream.
            generator.Save(stream, BarCodeImageFormat.Png);
        }
        stream.Position = 0;
        return stream;
    }

    // Legacy method delegates to the main implementation.
    public Stream GetOldBarcodeImage(Aspose.Words.Fields.BarcodeParameters parameters)
    {
        return GetBarcodeImage(parameters);
    }
}

public class Program
{
    public static void Main()
    {
        const string inputPath = "input.docx";
        const string outputPath = "output.pdf";

        // -----------------------------------------------------------------
        // Step 1: Create a sample DOCX containing a DISPLAYBARCODE field.
        // -----------------------------------------------------------------
        Document docCreate = new Document();
        DocumentBuilder builder = new DocumentBuilder(docCreate);

        // Insert a typed DISPLAYBARCODE field.
        FieldDisplayBarcode barcodeField = (FieldDisplayBarcode)builder.InsertField(FieldType.FieldDisplayBarcode, true);
        barcodeField.BarcodeType = "CODE128";
        barcodeField.BarcodeValue = "1234567890";
        barcodeField.ForegroundColor = "0x000000"; // Black
        barcodeField.BackgroundColor = "0xFFFFFF"; // White

        docCreate.UpdateFields();
        docCreate.Save(inputPath);

        // -----------------------------------------------------------------
        // Step 2: Load the existing DOCX, register the custom generator, and export to PDF.
        // -----------------------------------------------------------------
        Document doc = new Document(inputPath);

        // Register custom barcode generator for rendered output (PDF).
        doc.FieldOptions.BarcodeGenerator = new CustomBarcodeGenerator();

        // Render barcodes using the custom generator.
        doc.UpdateFields();

        // Save as PDF.
        doc.Save(outputPath);
    }
}
