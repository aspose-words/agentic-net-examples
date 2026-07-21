using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Fields;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Simulate barcode data that might be missing.
        string barcodeData = null; // Change to a non‑null value to see a valid barcode.

        // Insert a DISPLAYBARCODE field using the typed API.
        FieldDisplayBarcode barcodeField = (FieldDisplayBarcode)builder.InsertField(FieldType.FieldDisplayBarcode, true);

        // Set common barcode properties.
        barcodeField.BarcodeType = "QR";

        // Apply error handling for missing barcode data.
        if (string.IsNullOrWhiteSpace(barcodeData))
        {
            // If data is missing, set a placeholder value and lock the field to prevent update errors.
            barcodeField.BarcodeValue = "N/A";
            barcodeField.IsLocked = true;
        }
        else
        {
            barcodeField.BarcodeValue = barcodeData;
        }

        // Optional visual settings.
        barcodeField.BackgroundColor = "0xFFFFFF";
        barcodeField.ForegroundColor = "0x000000";
        barcodeField.ErrorCorrectionLevel = "3";
        barcodeField.ScalingFactor = "250";
        barcodeField.SymbolHeight = "1000";
        barcodeField.SymbolRotation = "0";

        // Update fields safely.
        try
        {
            doc.UpdateFields();
        }
        catch (Exception ex)
        {
            // Log the error and continue; the document will still be saved.
            Console.WriteLine("Field update failed: " + ex.Message);
        }

        // Save the document to the output folder.
        string outputPath = Path.Combine(Environment.CurrentDirectory, "BarcodeDocument.docx");
        doc.Save(outputPath);
    }
}
