using System;
using Aspose.Words;
using Aspose.Words.Tables;

class InsertDocumentIntoTable
{
    static void Main()
    {
        // Load the source DOCX that will be inserted.
        Document srcDoc = new Document("SourceDocument.docx");

        // Create a new blank destination document.
        Document dstDoc = new Document();

        // Initialize a DocumentBuilder for the destination document.
        DocumentBuilder builder = new DocumentBuilder(dstDoc);

        // Start a table with a single cell.
        builder.StartTable();
        builder.InsertCell();

        // Insert the source document into the current cell.
        // KeepSourceFormatting preserves the original styles of the source.
        builder.InsertDocument(srcDoc, ImportFormatMode.KeepSourceFormatting);

        // Close the row and the table.
        builder.EndRow();
        builder.EndTable();

        // Save the resulting document.
        dstDoc.Save("ResultDocument.docx");
    }
}
