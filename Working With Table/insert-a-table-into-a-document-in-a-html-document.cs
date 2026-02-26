using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Tables;

class Program
{
    static void Main()
    {
        // Create a new empty document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert an HTML fragment that contains a table.
        string htmlTable = @"
<table border='1' style='border-collapse:collapse;'>
    <tr><th>Header 1</th><th>Header 2</th></tr>
    <tr><td>Cell 1</td><td>Cell 2</td></tr>
    <tr><td>Cell 3</td><td>Cell 4</td></tr>
</table>";

        builder.InsertHtml(htmlTable);

        // Add a line break after the HTML table.
        builder.Writeln();

        // Insert an additional table using the DocumentBuilder API.
        Table table = builder.StartTable();
        builder.InsertCell();
        builder.Write("Row 1, Cell 1");
        builder.InsertCell();
        builder.Write("Row 1, Cell 2");
        builder.EndRow();

        builder.InsertCell();
        builder.Write("Row 2, Cell 1");
        builder.InsertCell();
        builder.Write("Row 2, Cell 2");
        builder.EndRow();

        builder.EndTable();

        // Save the document to the current directory.
        string outputPath = Path.Combine(Environment.CurrentDirectory, "Result.docx");
        doc.Save(outputPath);
    }
}
