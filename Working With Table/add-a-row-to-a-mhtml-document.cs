using System;
using Aspose.Words;
using Aspose.Words.Tables;
using Aspose.Words.Saving;

class Program
{
    static void Main()
    {
        // Load the existing MHTML document.
        Document doc = new Document("Input.mht");

        // Get the first table in the document (assumes at least one table exists).
        Table table = doc.FirstSection.Body.Tables[0];

        // Create a new row for the document.
        Row newRow = new Row(doc);

        // Create a new cell with some sample text.
        Cell cell = new Cell(doc);
        Paragraph paragraph = new Paragraph(doc);
        Run run = new Run(doc, "Added row cell");
        paragraph.AppendChild(run);
        cell.AppendChild(paragraph);

        // Add the cell to the newly created row.
        newRow.Cells.Add(cell);

        // Append the new row to the table.
        table.Rows.Add(newRow);

        // Save the modified document back to MHTML format.
        HtmlSaveOptions saveOptions = new HtmlSaveOptions(SaveFormat.Mhtml);
        doc.Save("Output.mht", saveOptions);
    }
}
