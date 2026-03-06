using System;
using Aspose.Words;
using Aspose.Words.Tables;

class InsertTextIntoCell
{
    static void Main()
    {
        // Path to the source DOCM file.
        string inputPath = @"C:\Docs\SourceDocument.docm";

        // Path where the modified DOCM will be saved.
        string outputPath = @"C:\Docs\ModifiedDocument.docm";

        // Load the existing DOCM document.
        Document doc = new Document(inputPath);

        // Create a DocumentBuilder attached to the loaded document.
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Move the builder's cursor to the desired cell.
        // Parameters: row index, column index, row span, column span (all zero‑based).
        // Example moves to the cell at first row, second column.
        builder.MoveToCell(0, 1, 0, 0);

        // Insert the desired text into the current cell.
        builder.Write("Inserted text into the cell.");

        // Save the document back as a DOCM file.
        doc.Save(outputPath);
    }
}
