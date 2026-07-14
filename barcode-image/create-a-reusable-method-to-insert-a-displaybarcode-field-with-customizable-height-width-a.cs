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

        // Insert a DISPLAYBARCODE field with custom parameters.
        InsertDisplayBarcode(builder, "123456789012", "EAN13", "1000", "250");

        // Update fields so the barcode is rendered in the document.
        doc.UpdateFields();

        // Save the document to the local file system.
        doc.Save("DisplayBarcode.docx");
    }

    // Reusable method that inserts a DISPLAYBARCODE field with customizable height and scaling.
    public static void InsertDisplayBarcode(DocumentBuilder builder, string barcodeValue, string barcodeType, string symbolHeight, string scalingFactor)
    {
        // Insert a typed DISPLAYBARCODE field.
        FieldDisplayBarcode field = (FieldDisplayBarcode)builder.InsertField(FieldType.FieldDisplayBarcode, true);

        // Set the required properties.
        field.BarcodeValue = barcodeValue;
        field.BarcodeType = barcodeType;
        field.SymbolHeight = symbolHeight;   // Height in TWIPS (1/1440 inch).
        field.ScalingFactor = scalingFactor; // Scaling factor as a percentage (10‑1000).
    }
}
