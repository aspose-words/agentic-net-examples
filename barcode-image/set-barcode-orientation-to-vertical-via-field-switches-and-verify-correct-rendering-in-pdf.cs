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
        // Generate barcode image
        using (var generator = new BarcodeGenerator(EncodeTypes.Code128, parameters.BarcodeValue))
        {
            // Set barcode colors
            generator.Parameters.Barcode.BarColor = Color.Black;
            generator.Parameters.BackColor = Color.White;

            using (var originalStream = new MemoryStream())
            {
                generator.Save(originalStream, BarCodeImageFormat.Png);
                originalStream.Position = 0;

                // Load the image, rotate it 90 degrees to achieve vertical orientation
                using (Image image = Image.FromStream(originalStream))
                {
                    image.RotateFlip(RotateFlipType.Rotate90FlipNone);
                    var rotatedStream = new MemoryStream();
                    image.Save(rotatedStream, ImageFormat.Png);
                    rotatedStream.Position = 0;
                    return rotatedStream;
                }
            }
        }
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
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert a DISPLAYBARCODE field
        FieldDisplayBarcode field = (FieldDisplayBarcode)builder.InsertField(FieldType.FieldDisplayBarcode, true);
        field.BarcodeType = "CODE128";
        field.BarcodeValue = "1234567890";

        // Set field colors (optional)
        field.ForegroundColor = "0x000000";
        field.BackgroundColor = "0xFFFFFF";

        // Register the custom barcode generator for rendered output (PDF)
        doc.FieldOptions.BarcodeGenerator = new CustomBarcodeGenerator();

        // Update fields to render the barcode
        doc.UpdateFields();

        // Save the document as PDF
        string outputPath = "BarcodeVertical.pdf";
        doc.Save(outputPath);

        // Simple verification: ensure the file was created and has content
        FileInfo info = new FileInfo(outputPath);
        Console.WriteLine($"PDF generated: {info.Exists && info.Length > 0}");
    }
}
