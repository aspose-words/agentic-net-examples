using Aspose.Words;
using Aspose.Words.Tables;

class Program
{
    static void Main()
    {
        // Create a new empty document.
        Document doc = new Document();

        // Use DocumentBuilder to construct the table.
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Begin a table and add two cells – one with English text, one with Arabic text.
        Table table = builder.StartTable();
        builder.InsertCell();
        builder.Write("Name");
        builder.InsertCell();
        builder.Write("مرحبًا"); // Arabic greeting.
        builder.EndRow();
        builder.EndTable();

        // Set the table to be right‑to‑left.
        table.Bidi = true;

        // Save the document to a file.
        string artifactsDir = @"C:\Temp\";
        doc.Save(artifactsDir + "RightToLeftTable.docx");
    }
}
