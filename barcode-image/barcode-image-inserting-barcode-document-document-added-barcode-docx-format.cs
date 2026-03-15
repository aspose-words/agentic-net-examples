using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Fields;

namespace AsposeWordsBarcodeDemo
{
    class Program
    {
        static void Main(string[] args)
        {
            // Path where the generated documents will be saved.
            string outputDir = Path.Combine(Environment.CurrentDirectory, "Output");
            Directory.CreateDirectory(outputDir);

            // 1. Create a new blank document and insert a DISPLAYBARCODE field (QR code).
            Document doc = new Document();                     // Create a blank document.
            DocumentBuilder builder = new DocumentBuilder(doc); // Helper to add content.

            // Insert a DISPLAYBARCODE field and configure it as a QR code with custom colors.
            FieldDisplayBarcode displayBarcode = (FieldDisplayBarcode)builder.InsertField(FieldType.FieldDisplayBarcode, true);
            displayBarcode.BarcodeType = "QR";
            displayBarcode.BarcodeValue = "ABC123";
            displayBarcode.BackgroundColor = "0xF8BD69";
            displayBarcode.ForegroundColor = "0xB5413B";
            displayBarcode.ErrorCorrectionLevel = "3";
            displayBarcode.ScalingFactor = "250";
            displayBarcode.SymbolHeight = "1000";
            displayBarcode.SymbolRotation = "0";

            // Add a paragraph break after the field for readability.
            builder.Writeln();

            // Save the document containing the barcode.
            string createdDocPath = Path.Combine(outputDir, "BarcodeCreated.docx");
            doc.Save(createdDocPath); // Save as DOCX (extension determines format).

            // 2. Load the previously saved document and add another barcode (EAN13) to it.
            Document loadedDoc = new Document(createdDocPath); // Load existing DOCX.
            DocumentBuilder loadedBuilder = new DocumentBuilder(loadedDoc);

            // Move the cursor to the end of the document.
            loadedBuilder.MoveToDocumentEnd();

            // Insert a second DISPLAYBARCODE field, this time an EAN13 barcode.
            FieldDisplayBarcode ean13Barcode = (FieldDisplayBarcode)loadedBuilder.InsertField(FieldType.FieldDisplayBarcode, true);
            ean13Barcode.BarcodeType = "EAN13";
            ean13Barcode.BarcodeValue = "501234567890";
            ean13Barcode.DisplayText = true;      // Show numeric text below the bars.
            ean13Barcode.PosCodeStyle = "CASE";   // Use CASE style for point‑of‑sale barcode.
            ean13Barcode.FixCheckDigit = true;    // Ensure a valid check digit.

            // Save the updated document.
            string updatedDocPath = Path.Combine(outputDir, "BarcodeUpdated.docx");
            loadedDoc.Save(updatedDocPath);
        }
    }
}
