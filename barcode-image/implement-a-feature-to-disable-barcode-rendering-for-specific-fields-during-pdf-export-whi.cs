using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Fields;
using Aspose.BarCode.Generation;
using Aspose.Drawing;
using Aspose.Drawing.Imaging;

public sealed class CustomBarcodeGenerator : IBarcodeGenerator
{
    public Stream GetBarcodeImage(Aspose.Words.Fields.BarcodeParameters parameters)
    {
        // If the barcode value starts with the disable marker, return a blank PNG.
        if (parameters.BarcodeValue != null && parameters.BarcodeValue.StartsWith("DISABLE:", StringComparison.Ordinal))
        {
            // Create a 1x1 white pixel image.
            using (var bitmap = new Bitmap(1, 1))
            {
                using (var graphics = Graphics.FromImage(bitmap))
                {
                    graphics.Clear(Color.White);
                }

                var blankStream = new MemoryStream();
                bitmap.Save(blankStream, ImageFormat.Png);
                blankStream.Position = 0;
                return blankStream;
            }
        }

        // Normal barcode generation.
        var stream = new MemoryStream();
        using (var generator = new BarcodeGenerator(EncodeTypes.Code128, parameters.BarcodeValue))
        {
            generator.Parameters.Barcode.BarColor = Aspose.Drawing.Color.Black;
            generator.Parameters.BackColor = Aspose.Drawing.Color.White;
            generator.Save(stream, BarCodeImageFormat.Png);
        }
        stream.Position = 0;
        return stream;
    }

    public Stream GetOldBarcodeImage(Aspose.Words.Fields.BarcodeParameters parameters)
    {
        // Delegate to the primary method.
        return GetBarcodeImage(parameters);
    }
}

public class Program
{
    public static void Main()
    {
        // Create a new document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert a normal barcode field.
        FieldDisplayBarcode normalField = (FieldDisplayBarcode)builder.InsertField(FieldType.FieldDisplayBarcode, true);
        normalField.BarcodeType = "CODE128";
        normalField.BarcodeValue = "1234567890";
        normalField.ForegroundColor = "0x000000";
        normalField.BackgroundColor = "0xFFFFFF";

        // Insert a barcode field that should be disabled in PDF.
        FieldDisplayBarcode disabledField = (FieldDisplayBarcode)builder.InsertField(FieldType.FieldDisplayBarcode, true);
        disabledField.BarcodeType = "CODE128";
        disabledField.BarcodeValue = "DISABLE:1234567890";
        disabledField.ForegroundColor = "0x000000";
        disabledField.BackgroundColor = "0xFFFFFF";

        // Update fields and save as DOCX (barcode fields remain as Word fields).
        doc.UpdateFields();
        doc.Save("output.docx");

        // Register the custom barcode generator for rendered outputs (PDF).
        doc.FieldOptions.BarcodeGenerator = new CustomBarcodeGenerator();

        // Update fields again to render barcodes using the custom generator.
        doc.UpdateFields();

        // Save the document as PDF; the disabled field will render as a blank image.
        doc.Save("output.pdf");
    }
}
