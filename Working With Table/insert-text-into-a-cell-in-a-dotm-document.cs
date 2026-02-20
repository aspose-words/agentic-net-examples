using System;
using Aspose.Words;
using Aspose.Words.Tables;

class InsertTextIntoCell
{
    static void Main()
    {
        // Path to the DOTM template.
        string inputPath = @"C:\Docs\Template.dotm";

        // Load the DOTM document.
        Document doc = new Document(inputPath);

        // Create a DocumentBuilder attached to the loaded document.
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Locate the first table in the document.
        Table table = doc.GetChild(NodeType.Table, 0, true) as Table;
        if (table == null)
        {
            Console.WriteLine("No table found in the document.");
            return;
        }

        // Choose the cell where the text will be inserted.
        // Example: first cell of the first row.
        Cell targetCell = table.Rows[0].Cells[0];

        // Move the builder's cursor to the beginning of the cell's first paragraph.
        builder.MoveTo(targetCell.FirstParagraph);

        // Insert the desired text.
        builder.Write("Inserted text");

        // Save the modified document as a macro‑enabled Word file.
        string outputPath = @"C:\Docs\Result.docm";
        doc.Save(outputPath);
    }
}
