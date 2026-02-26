using Aspose.Words;
using Aspose.Words.Tables;

class Program
{
    static void Main()
    {
        // Load the existing DOCM file.
        Document doc = new Document("Input.docm");

        // Attach a DocumentBuilder to the loaded document.
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Move the builder's cursor to the target cell.
        // Parameters: row index, column index, row span, column span (all zero‑based).
        builder.MoveToCell(0, 0, 0, 0);

        // Insert the desired text into the cell.
        builder.Write("Hello from Aspose.Words!");

        // Save the modified document.
        doc.Save("Output.docm");
    }
}
