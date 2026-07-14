using System;
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
        // Example: 2000 twips ≈ 1.39 inches.
        barcodeField.SymbolHeight = "2000";

        // Set the scaling factor (percentage) to control the width proportionally.
        // Example: 150% makes the barcode 1.5 times wider than the default.
        barcodeField.ScalingFactor = "150";

        // Update fields so that the changes are applied.
        doc.UpdateFields();

        // Save the document to the current directory.
        doc.Save("DisplayBarcodeHeightWidth.docx");
    }
}
