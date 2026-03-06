using Aspose.Words;
using Aspose.Words.Tables;

class Program
{
    static void Main()
    {
        // Load the existing WORDML (or DOCX) document.
        Document doc = new Document("Input.docx");

        // Create a DocumentBuilder to edit the document.
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Move the cursor to the target cell.
        // Parameters: tableIndex, rowIndex, columnIndex, characterIndex.
        // Here we target the first table (0), second row (1), third column (2),
        // and position the cursor at the end of the cell (-1).
        builder.MoveToCell(0, 1, 2, -1);

        // Insert the desired text into the cell.
        builder.Write("Inserted text into the cell.");

        // Save the modified document.
        doc.Save("Output.docx");
    }
}
