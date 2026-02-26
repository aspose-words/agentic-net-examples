using System;
using Aspose.Words;
using Aspose.Words.Tables;

class InsertHtmlIntoCell
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

        // Insert HTML content into the current cell.
        // The HTML fragment will be parsed and converted to Word formatting.
        builder.InsertHtml("<p align='center'><b>Hello, <span style='color:red;'>World</span>!</b></p>");

        // End the row and the table.
        builder.EndRow();
        builder.EndTable();

        // Save the document to disk.
        doc.Save("CellWithHtml.docx");
    }
}
