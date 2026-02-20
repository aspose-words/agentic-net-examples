using System;
using Aspose.Words;
using Aspose.Words.Tables;

class Program
{
    static void Main()
    {
        // HTML containing a simple table.
        const string html = @"
            <html>
                <body>
                    <table>
                        <tr><td>Cell 1</td><td>Cell 2</td></tr>
                        <tr><td>Cell 3</td><td>Cell 4</td></tr>
                    </table>
                </body>
            </html>";

        // Load the HTML into a new document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.InsertHtml(html);

        // Retrieve the first table in the document.
        Table table = (Table)doc.GetChild(NodeType.Table, 0, true);

        // Apply a built‑in table style (optional, helps see the effect).
        table.StyleIdentifier = StyleIdentifier.LightGrid;

        // Apply TableStyleOptions flags: first row, first column, and row banding.
        table.StyleOptions = TableStyleOptions.FirstRow |
                             TableStyleOptions.FirstColumn |
                             TableStyleOptions.RowBands;

        // Save the resulting document.
        doc.Save("StyledTable.docx");
    }
}
