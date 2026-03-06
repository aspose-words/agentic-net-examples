using System;
using Aspose.Words;
using Aspose.Words.Tables;

class Program
{
    static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();

        // Associate a DocumentBuilder with the document.
        DocumentBuilder builder = new DocumentBuilder(doc);

        // HTML fragment that defines a simple table.
        string htmlTable = @"
<table border='1' style='border-collapse:collapse;'>
    <tr>
        <th>Header 1</th>
        <th>Header 2</th>
    </tr>
    <tr>
        <td>Cell 1A</td>
        <td>Cell 1B</td>
    </tr>
    <tr>
        <td>Cell 2A</td>
        <td>Cell 2B</td>
    </tr>
</table>";

        // Insert the HTML (including the table) into the document.
        builder.InsertHtml(htmlTable);

        // Save the resulting document.
        doc.Save("TableFromHtml.docx");
    }
}
