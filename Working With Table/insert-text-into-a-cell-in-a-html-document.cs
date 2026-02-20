using Aspose.Words;
using Aspose.Words.Tables;

class Program
{
    static void Main()
    {
        // Load the HTML document.
        Document doc = new Document("input.html");

        // Locate the first table and the first cell within it.
        Table table = doc.FirstSection.Body.Tables[0];
        Cell cell = table.Rows[0].Cells[0];

        // Position the DocumentBuilder at the start of the cell's first paragraph.
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.MoveTo(cell.FirstParagraph);

        // Insert the desired text into the cell.
        builder.Write("Inserted text");

        // Save the updated document.
        doc.Save("output.docx");
    }
}
