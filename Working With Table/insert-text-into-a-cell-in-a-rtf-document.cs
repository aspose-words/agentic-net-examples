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

        // Start a table. The first call to InsertCell will also start the first row.
        builder.StartTable();

        // Insert the first cell of the table.
        builder.InsertCell();

        // Write the desired text into the current cell.
        builder.Write("This text is inside a table cell in an RTF document.");

        // End the row and the table.
        builder.EndRow();
        builder.EndTable();

        // Save the document as RTF. The file extension determines the format.
        doc.Save("CellText.rtf");
    }
}
