using System;
using Aspose.Words;
using Aspose.Words.BuildingBlocks;
using Aspose.Words.Fields;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert a DISPLAYBARCODE field with a placeholder value.
        // Use the typed insertion method as required by the rules.
        FieldDisplayBarcode barcodeField = (FieldDisplayBarcode)builder.InsertField(FieldType.FieldDisplayBarcode, true);
        barcodeField.BarcodeType = "QR";               // QR code type.
        barcodeField.BarcodeValue = "PLACEHOLDER";     // Placeholder that will be replaced later.
        barcodeField.BackgroundColor = "0xF8BD69";     // Optional styling.
        barcodeField.ForegroundColor = "0xB5413B";
        barcodeField.ErrorCorrectionLevel = "3";
        barcodeField.ScalingFactor = "250";
        barcodeField.SymbolHeight = "1000";
        barcodeField.SymbolRotation = "0";

        // Add a line break after the field for readability.
        builder.Writeln();

        // Replace the placeholder with the actual dynamic value.
        barcodeField.BarcodeValue = "https://example.com";

        // Update fields so the barcode image is generated with the new value.
        doc.UpdateFields();

        // Save the document to disk.
        doc.Save("DisplayBarcodeDynamic.docx");
    }
}
