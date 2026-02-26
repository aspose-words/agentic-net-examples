using System;
using Aspose.Words;
using Aspose.Words.Tables;

class Program
{
    static void Main()
    {
        // Load the DOTX template.
        Document doc = new Document("Template.dotx");

        // Create a DocumentBuilder for editing.
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Move the cursor to the first cell of the first table (table index 0, row 0, column 0, cell 0).
        builder.MoveToCell(0, 0, 0, 0);

        // Insert the desired text into the cell.
        builder.Write("Inserted text into the cell.");

        // Save the modified document.
        doc.Save("Result.docx");
    }
}
