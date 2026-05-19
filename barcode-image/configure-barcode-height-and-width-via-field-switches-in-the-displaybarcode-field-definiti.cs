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

        // Set basic barcode properties.
        barcodeField.BarcodeType = "QR";               // QR code type.
        barcodeField.BarcodeValue = "1234567890";      // Data to encode.

        // Configure the barcode's height (in TWIPS) and width scaling factor.
        // Height: 1500 TWIPS ≈ 1.04 inches.
        // ScalingFactor: 200 means the barcode width will be doubled.
        barcodeField.SymbolHeight = "1500";
        barcodeField.ScalingFactor = "200";

        // Apply the changes to the document.
        doc.UpdateFields();

        // Save the document to the local file system.
        doc.Save("DisplayBarcodeField.docx");
    }
}
