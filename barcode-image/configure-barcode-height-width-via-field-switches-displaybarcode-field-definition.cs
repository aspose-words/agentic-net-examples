using System;
using Aspose.Words;
using Aspose.Words.Fields;

class Program
{
    static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert a DISPLAYBARCODE field that will show a QR code.
        FieldDisplayBarcode field = (FieldDisplayBarcode)builder.InsertField(FieldType.FieldDisplayBarcode, true);

        // Basic barcode settings.
        field.BarcodeType = "QR";
        field.BarcodeValue = "ABC123";

        // Configure the symbol height (in TWIPS; 1 TWIP = 1/1440 inch).
        // Example: 1500 TWIPS ≈ 1.04 inches.
        field.SymbolHeight = "1500";

        // Configure the symbol width via scaling factor (percentage).
        // Example: 300 means the barcode will be 3 times wider than its default size.
        field.ScalingFactor = "300";

        // Optional visual tweaks.
        field.BackgroundColor = "0xF8BD69";
        field.ForegroundColor = "0xB5413B";

        // Add a line break after the field for readability.
        builder.Writeln();

        // Save the document to disk.
        doc.Save("DisplayBarcodeHeightWidth.docx");
    }
}
