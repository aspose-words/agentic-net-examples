using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Fields;
using Aspose.BarCode.Generation;
using Aspose.Drawing;

public sealed class CustomBarcodeGenerator : IBarcodeGenerator
{
    // Generates a barcode image and logs the generation event.
    public Stream GetBarcodeImage(Aspose.Words.Fields.BarcodeParameters parameters)
    {
        // Create a memory stream to hold the PNG image.
        MemoryStream stream = new MemoryStream();

        // Generate the barcode using Aspose.BarCode.
        using (BarcodeGenerator generator = new BarcodeGenerator(EncodeTypes.Code128, parameters.BarcodeValue))
        {
            // Set barcode colors.
            generator.Parameters.Barcode.BarColor = Aspose.Drawing.Color.Black;
            generator.Parameters.BackColor = Aspose.Drawing.Color.White;

            // Save the barcode as PNG into the stream.
            generator.Save(stream, BarCodeImageFormat.Png);
        }

        // Determine image dimensions.
        int width, height;
        stream.Position = 0;
        using (Image image = Image.FromStream(stream))
        {
            width = image.Width;
            height = image.Height;
        }

        // Log the generation event (using barcode value as identifier).
        using (StreamWriter log = new StreamWriter("barcode_log.txt", true))
        {
            string identifier = parameters.BarcodeValue ?? "<unknown>";
            log.WriteLine($"{DateTime.Now:u}: Barcode '{identifier}' generated image size {width}x{height}");
        }

        // Reset stream position for the caller.
        stream.Position = 0;
        return stream;
    }

    // Legacy method delegates to the primary implementation.
    public Stream GetOldBarcodeImage(Aspose.Words.Fields.BarcodeParameters parameters)
    {
        return GetBarcodeImage(parameters);
    }
}

public class Program
{
    public static void Main()
    {
        // Create a new document and a builder.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert a DISPLAYBARCODE field.
        FieldDisplayBarcode field = (FieldDisplayBarcode)builder.InsertField(FieldType.FieldDisplayBarcode, true);
        field.BarcodeType = "CODE128";
        field.BarcodeValue = "1234567890";

        // Register the custom barcode generator for rendered output (PDF).
        doc.FieldOptions.BarcodeGenerator = new CustomBarcodeGenerator();

        // Update fields to trigger barcode generation.
        doc.UpdateFields();

        // Save the document as PDF.
        doc.Save("output.pdf");
    }
}
