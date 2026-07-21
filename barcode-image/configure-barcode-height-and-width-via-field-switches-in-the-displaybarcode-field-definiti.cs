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
        barcodeField.BarcodeType = "CODE39";
        barcodeField.BarcodeValue = "12345ABCDE";

        // Set the height of the barcode symbol (in twips; 1 inch = 1440 twips).
        // Example: 2 inches high => 2 * 1440 = 2880 twips.
        barcodeField.SymbolHeight = "2880";

        // Set the scaling factor for the barcode width (percentage).
        // Example: 200% width.
        barcodeField.ScalingFactor = "200";

        // Update fields so that the field code reflects the set properties.
        doc.UpdateFields();

        // Save the document to the current directory.
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "BarcodeDisplay.docx");
        doc.Save(outputPath);
    }
}
