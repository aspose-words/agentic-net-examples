using System;
using Aspose.Words;
using Aspose.Words.Fields;

class InsertBarcodeOnEachPage
{
    static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Add some sample content that spans multiple pages.
        builder.Writeln("First page content.");
        builder.InsertBreak(BreakType.PageBreak);
        builder.Writeln("Second page content.");
        builder.InsertBreak(BreakType.PageBreak);
        builder.Writeln("Third page content.");

        // Move the builder to the primary header of the document.
        // The header is shared by all pages in the section, so the barcode will appear on every page.
        builder.MoveToHeaderFooter(HeaderFooterType.HeaderPrimary);

        // Insert a DISPLAYBARCODE field and configure its properties.
        // This field renders the barcode as an image when the document is opened in Word.
        FieldDisplayBarcode barcodeField = (FieldDisplayBarcode)builder.InsertField(FieldType.FieldDisplayBarcode, true);
        barcodeField.BarcodeType = "CODE39";          // Choose the barcode type.
        barcodeField.BarcodeValue = "12345ABCDE";    // The data to encode.
        barcodeField.AddStartStopChar = true;        // Add start/stop characters for CODE39.

        // Optionally, you can set additional visual properties such as colors, scaling, etc.
        // barcodeField.BackgroundColor = "0xFFFFFF";
        // barcodeField.ForegroundColor = "0x000000";

        // Save the document.
        doc.Save("BarcodeOnEachPage.docx");
    }
}
