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
        // The field initially contains a placeholder value that will be replaced later.
        FieldDisplayBarcode barcodeField = (FieldDisplayBarcode)builder.InsertField(FieldType.FieldDisplayBarcode, true);
        barcodeField.BarcodeType = "QR";               // Choose QR code for this example.
        barcodeField.BarcodeValue = "PLACEHOLDER";     // Placeholder text.
        barcodeField.BackgroundColor = "0xFFFFFF";     // White background.
        barcodeField.ForegroundColor = "0x000000";     // Black bars.

        // Simulate obtaining a dynamic value (e.g., from a database or user input).
        string dynamicValue = "ABC123";

        // Replace the placeholder with the dynamic value.
        barcodeField.BarcodeValue = dynamicValue;

        // Update all fields in the document so the DISPLAYBARCODE field reflects the new value.
        doc.UpdateFields();

        // Save the resulting document.
        doc.Save("DisplayBarcodeDynamic.docx");
    }
}
