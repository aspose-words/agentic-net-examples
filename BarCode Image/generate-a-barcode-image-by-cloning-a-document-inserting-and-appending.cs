using System;
using Aspose.Words;
using Aspose.Words.Fields;

class BarcodeDocumentProcessor
{
    static void Main()
    {
        // Load the original document.
        Document sourceDoc = new Document("Source.docx");

        // Clone the document (deep clone, including its content).
        Document clonedDoc = (Document)sourceDoc.Clone(true);

        // Prepare documents to be inserted and appended.
        Document insertDoc = new Document("Insert.docx");
        Document appendDoc = new Document("Append.docx");

        // Use DocumentBuilder to manipulate the cloned document.
        DocumentBuilder builder = new DocumentBuilder(clonedDoc);

        // Insert the insertDoc at the beginning of the cloned document.
        builder.MoveToDocumentStart();
        builder.InsertDocument(insertDoc, ImportFormatMode.KeepSourceFormatting);

        // Append the appendDoc at the end of the cloned document.
        builder.MoveToDocumentEnd();
        builder.InsertDocument(appendDoc, ImportFormatMode.KeepSourceFormatting);

        // Insert a DISPLAYBARCODE field (QR code) at the end of the document.
        // The field will generate the barcode image when fields are updated.
        builder.MoveToDocumentEnd();
        FieldBuilder barcodeBuilder = new FieldBuilder(FieldType.FieldDisplayBarcode);
        barcodeBuilder.AddArgument("QR");               // Barcode type.
        barcodeBuilder.AddArgument("ABC123");           // Barcode value.
        // Build and insert the field before a new empty paragraph.
        builder.Writeln(); // Ensure there is a paragraph to host the field.
        barcodeBuilder.BuildAndInsert(builder.CurrentParagraph);

        // Update all fields so that the barcode image is generated.
        clonedDoc.UpdateFields();

        // Save the modified cloned document.
        clonedDoc.Save("ClonedWithBarcode.docx");

        // Split the cloned document into separate documents, one per section.
        for (int sectionIndex = 0; sectionIndex < clonedDoc.Sections.Count; sectionIndex++)
        {
            // Clone the whole document to work on a copy.
            Document partDoc = (Document)clonedDoc.Clone(true);

            // Remove all sections except the current one.
            for (int i = partDoc.Sections.Count - 1; i >= 0; i--)
            {
                if (i != sectionIndex)
                    partDoc.Sections[i].Remove();
            }

            // Save the split part.
            string partFileName = $"ClonedPart_{sectionIndex + 1}.docx";
            partDoc.Save(partFileName);
        }
    }
}
