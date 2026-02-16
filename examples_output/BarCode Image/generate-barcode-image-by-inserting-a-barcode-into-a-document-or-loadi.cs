using System;
using Aspose.Words;
using Aspose.Words.Fields;

namespace BarcodeExample
{
    class Program
    {
        static void Main()
        {
            // Create a new blank document.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Insert a BARCODE field. The field syntax is:
            //   BARCODE <type> "<value>" [switches]
            // Here we insert a QR code with the value "ABC123".
            builder.InsertField("BARCODE QR \"ABC123\" ");

            // Save the document in DOCX format. When the document is opened in Word (or any
            // viewer that supports the BARCODE field) the barcode will be rendered.
            doc.Save("BarcodeDocument.docx");
        }
    }
}
