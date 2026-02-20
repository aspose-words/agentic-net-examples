using System;
using Aspose.Words;
using Aspose.Words.Tables;

class InsertTextIntoCell
{
    static void Main()
    {
        // Load an existing WORDML (or DOCX) document.
        Document doc = new Document("Input.docx");

        // Assume the document contains at least one table.
        Table table = doc.FirstSection.Body.Tables[0];

        // Choose the cell you want to modify.
        // For example, first row, second column (index is zero‑based).
        Cell targetCell = table.Rows[0].Cells[1];

        // Move the builder's cursor to the beginning of the cell's first paragraph.
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.MoveTo(targetCell.FirstParagraph);

        // Insert the desired text into the cell.
        builder.Write("Inserted text goes here.");

        // Save the modified document.
        doc.Save("Output.docx");
    }
}
