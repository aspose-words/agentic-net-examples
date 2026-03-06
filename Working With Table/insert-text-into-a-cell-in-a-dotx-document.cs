using System;
using Aspose.Words;
using Aspose.Words.Tables;

class Program
{
    static void Main()
    {
        // Load the DOTX template file.
        Document doc = new Document("Template.dotx");

        // Attach a DocumentBuilder to the loaded document.
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Move the builder's cursor to the first cell of the first table.
        // Parameters: tableIndex, rowIndex, columnIndex, cellIndex (all zero‑based).
        builder.MoveToCell(0, 0, 0, 0);

        // Insert the desired text into the selected cell.
        builder.Write("Hello from Aspose.Words!");

        // Save the modified document to a new file.
        doc.Save("Result.docx");
    }
}
