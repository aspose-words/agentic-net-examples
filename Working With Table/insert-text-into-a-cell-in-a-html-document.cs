using System;
using System.IO;
using System.Text;
using Aspose.Words;
using Aspose.Words.Tables;

class Program
{
    static void Main()
    {
        // Define a simple HTML string that contains a table.
        string html = @"
            <html>
                <body>
                    <table border='1'>
                        <tr><td>Cell 1</td><td>Cell 2</td></tr>
                        <tr><td>Cell 3</td><td>Cell 4</td></tr>
                    </table>
                </body>
            </html>";

        // Load the HTML into an Aspose.Words Document.
        Document doc;
        using (MemoryStream stream = new MemoryStream(Encoding.UTF8.GetBytes(html)))
        {
            doc = new Document(stream);
        }

        // Create a DocumentBuilder for editing the document.
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Move the cursor to the second cell of the first row (table index 0, row 0, column 1).
        // The characterIndex of 0 places the cursor at the start of the cell.
        builder.MoveToCell(0, 0, 1, 0);

        // Insert HTML content into the selected cell.
        builder.InsertHtml("<p><b>Inserted HTML</b> into cell.</p>");

        // If you prefer plain text instead of HTML, you could use:
        // builder.Write("Plain text inserted into cell.");

        // Save the modified document to disk.
        doc.Save("Result.docx");
    }
}
