using Aspose.Words;
using System;

class Program
{
    static void Main()
    {
        // Create a new empty document.
        Document doc = new Document();

        // Associate a DocumentBuilder with the document.
        DocumentBuilder builder = new DocumentBuilder(doc);

        // HTML fragment that defines a simple 2x2 table.
        string htmlTable = @"
            <table border='1' style='border-collapse:collapse;'>
                <tr>
                    <td>Row 1, Cell 1</td>
                    <td>Row 1, Cell 2</td>
                </tr>
                <tr>
                    <td>Row 2, Cell 1</td>
                    <td>Row 2, Cell 2</td>
                </tr>
            </table>";

        // Insert the HTML table into the document.
        builder.InsertHtml(htmlTable);

        // Save the resulting document.
        doc.Save("TableInHtml.docx");
    }
}
