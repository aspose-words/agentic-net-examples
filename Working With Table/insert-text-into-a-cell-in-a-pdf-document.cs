using System;
using Aspose.Words;
using Aspose.Words.Saving;

class Program
{
    static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();

        // Initialize a DocumentBuilder for the document.
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Build a simple 2x2 table.
        builder.StartTable();

        // First row.
        builder.InsertCell();               // Cell (0,0)
        builder.Write("Cell 0,0");
        builder.InsertCell();               // Cell (0,1)
        builder.Write("Cell 0,1");
        builder.EndRow();

        // Second row.
        builder.InsertCell();               // Cell (1,0)
        builder.Write("Cell 1,0");
        builder.InsertCell();               // Cell (1,1)
        builder.Write("Cell 1,1");
        builder.EndRow();

        // Finish the table.
        builder.EndTable();

        // Move the cursor to the second row, second column (row index 1, column index 1).
        builder.MoveToCell(tableIndex: 0, rowIndex: 1, columnIndex: 1, characterIndex: 0);

        // Insert the desired text into that cell.
        builder.Write("Inserted Text");

        // Save the document as a PDF file.
        PdfSaveOptions pdfOptions = new PdfSaveOptions();
        doc.Save("Output.pdf", pdfOptions);
    }
}
