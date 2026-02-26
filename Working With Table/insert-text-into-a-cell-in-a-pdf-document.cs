using System;
using Aspose.Words;
using Aspose.Words.Saving;
using Aspose.Words.Tables; // Added for Cell class

class InsertTextIntoPdfCell
{
    static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();

        // Initialize DocumentBuilder for the document.
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Start a table and add a single cell.
        builder.StartTable();
        Cell cell = builder.InsertCell();   // Insert a new cell.
        builder.Write("Hello, PDF cell!"); // Write text into the current cell.
        builder.EndRow();                  // End the row.
        builder.EndTable();                // End the table.

        // Save the document as PDF.
        PdfSaveOptions saveOptions = new PdfSaveOptions
        {
            // Optional: set text compression to Flate for smaller file size.
            TextCompression = PdfTextCompression.Flate
        };
        doc.Save("Output.pdf", saveOptions);
    }
}
