using System;
using Aspose.Words;
using Aspose.Words.Tables;
using Aspose.Words.Saving;

class Program
{
    static void Main()
    {
        // Load an existing Markdown document that already contains a table.
        Document doc = new Document("Input.md");

        // Assume the document has at least one table.
        Table table = doc.FirstSection.Body.Tables[0];

        // Choose the cell where the text should be inserted.
        // For example, insert into the second row, second column (index is zero‑based).
        Cell targetCell = table.Rows[1].Cells[1];

        // Remove any existing content from the cell.
        targetCell.RemoveAllChildren();

        // Create a new paragraph that contains the desired text.
        Paragraph paragraph = new Paragraph(doc);
        Run run = new Run(doc, "Inserted text");
        paragraph.AppendChild(run);
        targetCell.AppendChild(paragraph);

        // Save the modified document back to Markdown format.
        MarkdownSaveOptions saveOptions = new MarkdownSaveOptions();
        doc.Save("Output.md", saveOptions);
    }
}
