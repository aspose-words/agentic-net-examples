using System;
using Aspose.Words;
using Aspose.Words.Tables;

class Program
{
    static void Main()
    {
        // Paths to the source document (to be inserted) and the destination document (contains the table).
        string sourcePath = "Source.docx";
        string destinationPath = "Destination.docx";
        string outputPath = "Result.docx";

        // Load the destination document from disk.
        Document destinationDoc = new Document(destinationPath);

        // Create a DocumentBuilder attached to the destination document.
        DocumentBuilder builder = new DocumentBuilder(destinationDoc);

        // Ensure there is at least one table in the destination document.
        // If the document already contains a table, this block can be omitted.
        if (destinationDoc.FirstSection.Body.Tables.Count == 0)
        {
            // Start a new table with a single cell.
            builder.StartTable();
            builder.InsertCell(); // First (and only) cell.
            builder.EndRow();
            builder.EndTable();
        }

        // Move the builder's cursor to the first cell of the first table.
        // Parameters: sectionIndex, tableIndex, rowIndex, cellIndex (all zero‑based).
        builder.MoveToCell(0, 0, 0, 0);

        // Load the source document that will be inserted.
        Document sourceDoc = new Document(sourcePath);

        // Insert the source document at the current cursor position (inside the table cell).
        // KeepSourceFormatting preserves the original formatting of the inserted content.
        builder.InsertDocument(sourceDoc, ImportFormatMode.KeepSourceFormatting);

        // Save the modified destination document to a new file.
        destinationDoc.Save(outputPath);
    }
}
