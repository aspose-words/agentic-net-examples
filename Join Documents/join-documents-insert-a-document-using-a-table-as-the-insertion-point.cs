using System;
using Aspose.Words;

class JoinDocumentsWithTableInsertion
{
    static void Main()
    {
        // Paths to the source documents and the output file.
        string mainDocPath = "Main.docx";      // Document that contains the table.
        string insertDocPath = "Insert.docx";  // Document to be inserted.
        string outputPath = "Result.docx";

        // Load the main document (the one that has the table).
        Document mainDoc = new Document(mainDocPath);

        // Create a DocumentBuilder for the main document.
        DocumentBuilder builder = new DocumentBuilder(mainDoc);

        // Move the cursor to the desired cell in the first table.
        // Parameters: tableIndex, rowIndex, columnIndex, cellIndex.
        // Here we target the first cell (0,0) of the first table (index 0).
        builder.MoveToCell(0, 0, 0, 0);

        // Load the document that will be inserted.
        Document docToInsert = new Document(insertDocPath);

        // Insert the document at the current cursor position (inside the table cell).
        // KeepSourceFormatting preserves the original formatting of the inserted document.
        builder.InsertDocument(docToInsert, ImportFormatMode.KeepSourceFormatting);

        // Save the resulting document.
        mainDoc.Save(outputPath);
    }
}
