using System;
using Aspose.Words;
using Aspose.Words.Fields;

public class Program
{
    // Inserts a DISPLAYBARCODE field with customizable type, value, height and width.
    // Height and width are approximated via the ScalingFactor property (percentage).
    public static void InsertDisplayBarcode(DocumentBuilder builder, string barcodeValue, string barcodeType, int height, int width)
    {
        // Insert a typed DISPLAYBARCODE field.
        FieldDisplayBarcode field = (FieldDisplayBarcode)builder.InsertField(FieldType.FieldDisplayBarcode, true);

        // Set mandatory properties.
        field.BarcodeValue = barcodeValue;
        field.BarcodeType = barcodeType;

        // Approximate size using ScalingFactor (percentage as string).
        // Here we take the average of height and width as a simple scaling heuristic.
        int scaling = (height + width) / 2;
        if (scaling <= 0)
            scaling = 100; // default 100%

        field.ScalingFactor = scaling.ToString();

        // Optional visual settings.
        field.ForegroundColor = "0x000000"; // black
        field.BackgroundColor = "0xFFFFFF"; // white
    }

    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert a barcode: Code128, value "1234567890", size 150x150 (approx.).
        InsertDisplayBarcode(builder, "1234567890", "CODE128", 150, 150);

        // Update fields so the barcode is rendered in the document.
        doc.UpdateFields();

        // Save the document.
        doc.Save("output.docx");
    }
}
