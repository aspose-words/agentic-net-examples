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

        // Insert a DISPLAYBARCODE field using the typed API.
        FieldDisplayBarcode barcodeField = (FieldDisplayBarcode)builder.InsertField(FieldType.FieldDisplayBarcode, true);

        // Configure the barcode type and value.
        barcodeField.BarcodeType = "QR";
        barcodeField.BarcodeValue = "1234567890";

        // Set the height of the barcode symbol (in twips; 1 inch = 1440 twips).
        // Example: 2 inches high => 2 * 1440 = 2880 twips.
        barcodeField.SymbolHeight = "2880";

        // Set the width scaling factor as a percentage.
        // Example: 200% width.
        barcodeField.ScalingFactor = "200";

        // Apply the changes to the field.
        doc.UpdateFields();

        // Save the document to the current directory.
        string outputPath = Path.Combine(Environment.CurrentDirectory, "BarcodeDisplay.docx");
        doc.Save(outputPath);
    }
}
