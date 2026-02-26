using System;
using Aspose.Words;
using Aspose.Words.Tables;

class InsertTextIntoRtfCell
{
    static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();

        // Initialize a DocumentBuilder for the document.
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Start a table and insert the first cell.
        builder.StartTable();
        builder.InsertCell();

        // Write the desired text into the current cell.
        builder.Write("Hello, this text is inside an RTF table cell.");

        // End the table (optional if only one cell is needed).
        builder.EndTable();

        // Save the document as an RTF file.
        doc.Save("CellWithText.rtf");
    }
}
