using System;
using Aspose.Words;

namespace BarcodeExample
{
    class Program
    {
        static void Main()
        {
            // Create a new blank document.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Define barcode parameters (example: QR code).
            // These values correspond to the BARCODE field switches.
            string barcodeType = "QR";                 // \b switch – barcode type
            string barcodeValue = "ABC123";            // the data to encode
            string backgroundColor = "0xF8BD69";       // \b switch – background colour (hex)
            string foregroundColor = "0xB5413B";       // \f switch – foreground colour (hex)
            string errorCorrectionLevel = "3";        // \e switch – error correction level (for QR)
            string scalingFactor = "250";             // \s switch – scaling factor (percentage)
            string symbolHeight = "1000";             // \h switch – symbol height (in points)
            string symbolRotation = "0";              // \r switch – rotation angle (degrees)

            // Insert a BARCODE field with the above options.
            // The field syntax is:
            //   BARCODE <type> "<value>" \b <bg> \f <fg> \e <ec> \s <scale> \h <height> \r <rotation>
            string fieldCode = $"BARCODE {barcodeType} \"{barcodeValue}\" " +
                               $"\\b {backgroundColor} " +
                               $"\\f {foregroundColor} " +
                               $"\\e {errorCorrectionLevel} " +
                               $"\\s {scalingFactor} " +
                               $"\\h {symbolHeight} " +
                               $"\\r {symbolRotation}";

            builder.InsertField(fieldCode);

            // Update fields so the BARCODE field is rendered as an image.
            doc.UpdateFields();

            // Save the document to a DOCX file.
            doc.Save("BarcodeDocument.docx");
        }
    }
}
