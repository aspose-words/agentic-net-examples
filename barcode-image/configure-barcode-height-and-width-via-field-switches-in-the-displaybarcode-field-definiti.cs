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
        barcodeField.BarcodeValue = "Aspose";

        // Set the height of the barcode symbol (in TWIPS; 1 inch = 1440 TWIPS).
        // Example: 2 inches high.
        barcodeField.SymbolHeight = (2 * 1440).ToString();

        // Set the scaling factor to control the width (percentage).
        // Example: 150% scaling.
        barcodeField.ScalingFactor = "150";

        // Update fields to apply the changes.
        doc.UpdateFields();

        // Save the document to the current directory.
        string outputPath = Path.Combine(Environment.CurrentDirectory, "DisplayBarcode.docx");
        doc.Save(outputPath);
    }
}
