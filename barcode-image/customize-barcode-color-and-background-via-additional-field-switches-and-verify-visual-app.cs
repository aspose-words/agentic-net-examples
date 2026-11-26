using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Fields;
using Aspose.Words.Drawing;
using Aspose.BarCode.Generation;
using Aspose.Drawing;

public sealed class CustomBarcodeGenerator : IBarcodeGenerator
{
    // Generates a barcode image based on the parameters supplied by the DISPLAYBARCODE field.
    public Stream GetBarcodeImage(Aspose.Words.Fields.BarcodeParameters parameters)
    {
        // Stream that will hold the PNG image.
        MemoryStream stream = new MemoryStream();

        // Create the barcode generator with the requested symbology and value.
        using (BarcodeGenerator generator = new BarcodeGenerator(EncodeTypes.Code128, parameters.BarcodeValue))
        {
            // Apply foreground (barcode) color if provided.
            if (!string.IsNullOrEmpty(parameters.ForegroundColor))
            {
                // Convert a hex string like "0x0000FF" to an Aspose.Drawing.Color.
                string hex = parameters.ForegroundColor.StartsWith("0x", StringComparison.OrdinalIgnoreCase)
                    ? parameters.ForegroundColor.Substring(2)
                    : parameters.ForegroundColor;

                if (int.TryParse(hex, System.Globalization.NumberStyles.HexNumber, null, out int rgb))
                {
                    int r = (rgb >> 16) & 0xFF;
                    int g = (rgb >> 8) & 0xFF;
                    int b = rgb & 0xFF;
                    generator.Parameters.Barcode.BarColor = Color.FromArgb(r, g, b);
                }
            }

            // Apply background color if provided.
            if (!string.IsNullOrEmpty(parameters.BackgroundColor))
            {
                string hex = parameters.BackgroundColor.StartsWith("0x", StringComparison.OrdinalIgnoreCase)
                    ? parameters.BackgroundColor.Substring(2)
                    : parameters.BackgroundColor;

                if (int.TryParse(hex, System.Globalization.NumberStyles.HexNumber, null, out int rgb))
                {
                    int r = (rgb >> 16) & 0xFF;
                    int g = (rgb >> 8) & 0xFF;
                    int b = rgb & 0xFF;
                    generator.Parameters.BackColor = Color.FromArgb(r, g, b);
                }
            }

            // Save the barcode as PNG into the stream.
            generator.Save(stream, BarCodeImageFormat.Png);
        }

        // Reset the stream position for reading.
        stream.Position = 0;
        return stream;
    }

    // Legacy method – simply forwards to the primary implementation.
    public Stream GetOldBarcodeImage(Aspose.Words.Fields.BarcodeParameters parameters)
    {
        return GetBarcodeImage(parameters);
    }
}

public class Program
{
    public static void Main()
    {
        // Create a new empty document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert a DISPLAYBARCODE field using the typed API.
        FieldDisplayBarcode field = (FieldDisplayBarcode)builder.InsertField(FieldType.FieldDisplayBarcode, true);
        field.BarcodeType = "CODE128";
        field.BarcodeValue = "1234567890";

        // Set custom colors using hex strings (foreground = blue, background = yellow).
        field.ForegroundColor = "0x0000FF";
        field.BackgroundColor = "0xFFFF00";

        // Register the custom barcode generator for rendered output (PDF).
        doc.FieldOptions.BarcodeGenerator = new CustomBarcodeGenerator();

        // Update fields to render the barcode.
        doc.UpdateFields();

        // Verify that a shape (image) was created for the barcode.
        bool barcodeRendered = false;
        NodeCollection shapes = doc.GetChildNodes(NodeType.Shape, true);
        foreach (Shape shape in shapes)
        {
            if (shape.HasImage)
            {
                barcodeRendered = true;
                break;
            }
        }

        // Save the document as PDF.
        string outputPath = "output.pdf";
        doc.Save(outputPath);

        // Output verification result.
        Console.WriteLine(barcodeRendered
            ? "Barcode rendered successfully and saved to PDF."
            : "Barcode rendering failed.");
    }
}
