using System;
using Aspose.Words;
using Aspose.Words.Tables;

class Program
{
    static void Main()
    {
        // Load an existing Word document
        Document doc = new Document("Input.docx");

        // Get the first table in the document (adjust index as needed)
        Table table = doc.FirstSection.Body.Tables[0];

        // Create a new row belonging to the same document
        Row newRow = new Row(doc);
        // Ensure the row has at least one cell (Aspose.Words creates a default cell)
        newRow.EnsureMinimum();

        // Create a new cell with a paragraph and some text
        Cell newCell = new Cell(doc);
        Paragraph para = new Paragraph(doc);
        Run run = new Run(doc, "New cell content");
        para.AppendChild(run);
        newCell.AppendChild(para);

        // Add the cell to the new row
        newRow.AppendChild(newCell);

        // Append the new row to the end of the table
        table.AppendChild(newRow);

        // Save the modified document
        doc.Save("Output.docx");
    }
}
