using System;
using Aspose.Words;
using Aspose.Words.Saving;
using Aspose.Words.Tables; // Added for Table class

class RenderDocumentWithColumns
{
    static void Main()
    {
        // Number of columns to create in the document.
        int columnCount = 3;

        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Start a table that will act as a simple column layout.
        Table table = builder.StartTable();

        // Create the specified number of columns.
        for (int i = 0; i < columnCount; i++)
        {
            // Insert a new cell for each column.
            builder.InsertCell();

            // Add some placeholder text to the cell.
            builder.Writeln($"Column {i + 1}");
        }

        // End the row that contains the cells.
        builder.EndRow();

        // End the table.
        builder.EndTable();

        // Prepare PDF save options (default options are sufficient for this task).
        PdfSaveOptions pdfOptions = new PdfSaveOptions();

        // Save the document as a PDF file.
        doc.Save("RenderedDocument.pdf", pdfOptions);
    }
}
