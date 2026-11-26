using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Fields;
using Aspose.BarCode.Generation;
using Aspose.Words.Replacing;

public sealed class CustomBarcodeGenerator : IBarcodeGenerator
{
    public Stream GetBarcodeImage(Aspose.Words.Fields.BarcodeParameters parameters)
    {
        var stream = new MemoryStream();
        using (var generator = new BarcodeGenerator(EncodeTypes.Code128, parameters.BarcodeValue))
        {
            // Optional: configure colors if needed
            generator.Parameters.Barcode.BarColor = Aspose.Drawing.Color.Black;
            generator.Parameters.BackColor = Aspose.Drawing.Color.White;

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
        // Create a new document
        var doc = new Document();
        var builder = new DocumentBuilder(doc);

        // Insert a DISPLAYBARCODE field using the typed approach
        var field = (FieldDisplayBarcode)builder.InsertField(FieldType.FieldDisplayBarcode, true);
        field.BarcodeType = "CODE128";
        field.BarcodeValue = "1234567890";

        // Configure visual appearance
        field.ForegroundColor = "0x000000"; // Black
        field.BackgroundColor = "0xFFFFFF"; // White

        // Configure size via scaling factor (affects height and width)
        field.ScalingFactor = "200"; // 200% of default size

        // Register the custom barcode generator for rendered output (PDF)
        doc.FieldOptions.BarcodeGenerator = new CustomBarcodeGenerator();

        // Update fields to generate the barcode image
        doc.UpdateFields();

        // Save the document as PDF (rendered barcode)
        doc.Save("output.pdf");
    }
}
