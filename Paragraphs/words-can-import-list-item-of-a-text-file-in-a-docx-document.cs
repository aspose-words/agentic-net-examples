using System;
using Aspose.Words;
using Aspose.Words.Loading;

class ImportListFromText
{
    static void Main()
    {
        // Path to the plain‑text file that contains list items.
        string txtFilePath = @"C:\Data\list.txt";

        // Path where the resulting DOCX document will be saved.
        string outputDocxPath = @"C:\Data\Result.docx";

        // -----------------------------------------------------------------
        // Load the text file as a Word document.
        // TxtLoadOptions enables automatic detection of numbered/bulleted lists.
        // -----------------------------------------------------------------
        TxtLoadOptions loadOptions = new TxtLoadOptions
        {
            // Detect list items that use whitespace as a delimiter (e.g., "1 Item").
            DetectNumberingWithWhitespaces = true
        };
        Document sourceDoc = new Document(txtFilePath, loadOptions);

        // -----------------------------------------------------------------
        // Create a new blank document that will receive the imported list.
        // -----------------------------------------------------------------
        Document destDoc = new Document();

        // Use DocumentBuilder to position the cursor at the end of the document.
        DocumentBuilder builder = new DocumentBuilder(destDoc);
        builder.MoveToDocumentEnd();

        // Insert the source document (which now contains proper list formatting)
        // into the destination document. KeepSourceFormatting preserves the list
        // appearance from the source.
        builder.InsertDocument(sourceDoc, ImportFormatMode.KeepSourceFormatting);

        // -----------------------------------------------------------------
        // Save the combined document as DOCX.
        // -----------------------------------------------------------------
        destDoc.Save(outputDocxPath);
    }
}
