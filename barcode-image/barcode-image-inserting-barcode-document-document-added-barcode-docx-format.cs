using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Fields;

namespace BarcodeExample
{
    class Program
    {
        static void Main()
        {
            // -----------------------------------------------------------------
            // 1. Create a new document and insert a DISPLAYBARCODE field.
            // -----------------------------------------------------------------
            Document newDoc = new Document();                     // Create a blank document.
            DocumentBuilder builder = new DocumentBuilder(newDoc); // Helper to add content.

            // Insert a DISPLAYBARCODE field that will render a QR code.
            FieldDisplayBarcode displayField = (FieldDisplayBarcode)builder.InsertField(FieldType.FieldDisplayBarcode, true);
            displayField.BarcodeType = "QR";                     // Type of barcode.
            displayField.BarcodeValue = "ABC123";                // Data to encode.
            displayField.BackgroundColor = "0xF8BD69";           // Optional background colour.
            displayField.ForegroundColor = "0xB5413B";           // Optional foreground colour.
            displayField.ErrorCorrectionLevel = "3";             // QR error correction.
            displayField.ScalingFactor = "250";                  // Scale the symbol.
            displayField.SymbolHeight = "1000";                  // Height in TWIPS.
            displayField.SymbolRotation = "0";                   // No rotation.

            // Add a paragraph break after the field for readability.
            builder.Writeln();

            // Save the newly created document as DOCX.
            string newDocPath = Path.Combine(Environment.CurrentDirectory, "BarcodeInserted.docx");
            newDoc.Save(newDocPath);
            Console.WriteLine($"Created document saved to: {newDocPath}");

            // -----------------------------------------------------------------
            // 2. Load an existing document, insert another barcode, and save.
            // -----------------------------------------------------------------
            // Assume there is an existing DOCX file named "Template.docx" in the same folder.
            string templatePath = Path.Combine(Environment.CurrentDirectory, "Template.docx");
            if (!File.Exists(templatePath))
            {
                // If the template does not exist, create a simple placeholder file.
                Document placeholder = new Document();
                DocumentBuilder phBuilder = new DocumentBuilder(placeholder);
                phBuilder.Writeln("This is a placeholder template document.");
                placeholder.Save(templatePath);
            }

            // Load the existing document.
            Document loadedDoc = new Document(templatePath);
            DocumentBuilder loadBuilder = new DocumentBuilder(loadedDoc);

            // Move the cursor to the end of the document.
            loadBuilder.MoveToDocumentEnd();

            // Insert a DISPLAYBARCODE field that will render an EAN13 barcode.
            FieldDisplayBarcode eanField = (FieldDisplayBarcode)loadBuilder.InsertField(FieldType.FieldDisplayBarcode, true);
            eanField.BarcodeType = "EAN13";
            eanField.BarcodeValue = "501234567890";
            eanField.DisplayText = true;          // Show the numeric text below the bars.
            eanField.PosCodeStyle = "CASE";       // Point‑of‑sale style.
            eanField.FixCheckDigit = true;        // Ensure a valid check digit.

            // Save the modified document.
            string loadedDocPath = Path.Combine(Environment.CurrentDirectory, "BarcodeAddedToTemplate.docx");
            loadedDoc.Save(loadedDocPath);
            Console.WriteLine($"Loaded document saved to: {loadedDocPath}");
        }
    }
}
