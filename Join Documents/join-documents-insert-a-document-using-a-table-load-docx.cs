using System;
using Aspose.Words;
using Aspose.Words.Tables;

class Program
{
    static void Main()
    {
        // Load the document that will be inserted.
        Document srcDoc = new Document("Source.docx");

        // Create a new blank document that will hold the result.
        Document dstDoc = new Document();

        // Initialize a DocumentBuilder for the destination document.
        DocumentBuilder builder = new DocumentBuilder(dstDoc);

        // Start a table and add a single cell where the source document will be placed.
        Table table = builder.StartTable();
        builder.InsertCell();

        // Insert the source document into the current cell.
        // KeepSourceFormatting preserves the original formatting of the inserted content.
        builder.InsertDocument(srcDoc, ImportFormatMode.KeepSourceFormatting);

        // Close the row and the table.
        builder.EndRow();
        builder.EndTable();

        // Save the combined document.
        dstDoc.Save("Joined.docx");
    }
}
