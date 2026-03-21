using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Fields;

class Program
{
    static void Main()
    {
        const string templatePath = "Template.docx";

        // Ensure the template exists; if not, create a simple one with a merge field named "QRPlaceholder".
        Document doc;
        if (File.Exists(templatePath))
        {
            doc = new Document(templatePath);
        }
        else
        {
            doc = new Document();
            DocumentBuilder tempBuilder = new DocumentBuilder(doc);
            tempBuilder.Writeln("This is a sample document.");
            tempBuilder.InsertField("MERGEFIELD QRPlaceholder \\* MERGEFORMAT");
            doc.Save(templatePath);
        }

        DocumentBuilder builder = new DocumentBuilder(doc);

        // Assign a custom barcode generator.
        doc.FieldOptions.BarcodeGenerator = new CustomBarcodeGenerator();

        // Set up QR code parameters.
        BarcodeParameters parameters = new BarcodeParameters
        {
            BarcodeType = "QR",
            BarcodeValue = "https://example.com",
            BackgroundColor = "0xFFFFFF",
            ForegroundColor = "0x000000",
            ErrorCorrectionLevel = "3",
            ScalingFactor = "250",
            SymbolHeight = "1000",
            SymbolRotation = "0"
        };

        // Generate the QR code image and insert it at the placeholder position.
        using (Stream imgStream = doc.FieldOptions.BarcodeGenerator.GetBarcodeImage(parameters))
        {
            // Move the builder to the merge field that acts as the placeholder.
            builder.MoveToMergeField("QRPlaceholder", true, false);
            // Insert the generated QR code image.
            builder.InsertImage(imgStream);
        }

        // Save the document with the QR code image replacing the placeholder.
        doc.Save("Result.docx");
    }
}

// Minimal stub for a custom barcode generator.
// Returns a simple 1x1 PNG image.
public class CustomBarcodeGenerator : IBarcodeGenerator
{
    // Base64-encoded 1x1 black PNG.
    private static readonly byte[] PngBytes = Convert.FromBase64String(
        "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAwMCAO+XK6cAAAAASUVORK5CYII=");

    public Stream GetBarcodeImage(BarcodeParameters parameters)
    {
        return new MemoryStream(PngBytes) { Position = 0 };
    }

    public Stream GetOldBarcodeImage(BarcodeParameters parameters)
    {
        // Not required for this example; return the same placeholder image.
        return GetBarcodeImage(parameters);
    }
}
