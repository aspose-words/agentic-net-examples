using System;
using Aspose.Words;
using Aspose.Words.Fields;

public class Program
{
    // Inserts a DISPLAYBARCODE field with the specified parameters.
    // height and scalingFactor are strings because the field properties expect string values.
    public static void InsertDisplayBarcode(DocumentBuilder builder, string barcodeValue, string barcodeType, string symbolHeight, string scalingFactor)
    {
        // Insert a DISPLAYBARCODE field.
        FieldDisplayBarcode field = (FieldDisplayBarcode)builder.InsertField(FieldType.FieldDisplayBarcode, true);

        // Set required properties.
        field.BarcodeValue = barcodeValue;
        field.BarcodeType = barcodeType;
        field.SymbolHeight = symbolHeight;      // Height in TWIPS (1/1440 inch).
        field.ScalingFactor = scalingFactor;    // Scaling factor as a percentage (e.g., "250").

        // Update the document fields to reflect the changes.
        builder.Document.UpdateFields();
    }

    public static void Main()
    {
        // Create a new empty document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Example: Insert a QR code with custom size.
        InsertDisplayBarcode(builder, "ABC123", "QR", "1000", "250");

        // Example: Insert an EAN13 barcode with default size.
        InsertDisplayBarcode(builder, "501234567890", "EAN13", "800", "200");

        // Save the document to the local file system.
        doc.Save("DisplayBarcode.docx");
    }
}
