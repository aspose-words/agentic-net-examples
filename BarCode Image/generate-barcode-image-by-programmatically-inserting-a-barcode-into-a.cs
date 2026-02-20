using System;
using Aspose.Words;
using Aspose.Words.Fields;

namespace BarcodeFieldExample
{
    class Program
    {
        static void Main()
        {
            // Create a new blank document.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Insert a DISPLAYBARCODE field.
            // Syntax: DISPLAYBARCODE <BarcodeType> "<BarcodeValue>"
            // Example: DISPLAYBARCODE QR "ABC123"
            builder.Writeln("DisplayBarcode field:");
            builder.InsertField(@"DISPLAYBARCODE QR ""ABC123""");

            // Insert a blank line for readability.
            builder.Writeln();

            // Insert a MERGEBARCODE field.
            // Syntax: MERGEBARCODE <BarcodeType> "<BarcodeValue>"
            // Example: MERGEBARCODE CODE39 "12345ABCDE"
            builder.Writeln("MergeBarcode field:");
            builder.InsertField(@"MERGEBARCODE CODE39 ""12345ABCDE""");

            // Save the document in DOCX format.
            doc.Save("BarcodeFields.docx");
        }
    }
}
