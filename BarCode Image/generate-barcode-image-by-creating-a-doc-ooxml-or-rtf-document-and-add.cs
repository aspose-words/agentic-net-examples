using System;
using Aspose.Words;
using Aspose.Words.Fields;
using Aspose.Words.Saving;

class GenerateBarcodeDocument
{
    static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();

        // Ensure the document has at least one paragraph to insert the field into.
        Paragraph para = doc.FirstSection.Body.FirstParagraph ?? doc.FirstSection.Body.AppendParagraph(string.Empty);

        // Build a DISPLAYBARCODE field that will render a QR code with the value "ABC123".
        // The field code syntax: DISPLAYBARCODE \b QR \d ABC123
        FieldBuilder fieldBuilder = new FieldBuilder(FieldType.FieldDisplayBarcode);
        fieldBuilder.AddSwitch("\\b", "QR");          // Set the barcode type.
        fieldBuilder.AddSwitch("\\d", "ABC123");     // Set the barcode value.

        // Insert the field at the end of the paragraph.
        fieldBuilder.BuildAndInsert(para);

        // Save the document in DOCX format using OoxmlSaveOptions.
        OoxmlSaveOptions saveOptions = new OoxmlSaveOptions
        {
            // Example: enable pretty formatting for easier inspection (optional).
            PrettyFormat = true
        };
        doc.Save("BarcodeDocument.docx", saveOptions);
    }
}
