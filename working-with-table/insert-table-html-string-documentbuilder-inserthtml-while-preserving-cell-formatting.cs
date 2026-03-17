using System;
using Aspose.Words;
using Aspose.Words.Saving;

class InsertHtmlTableExample
{
    static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();

        // Attach a DocumentBuilder to the document.
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Set some cell formatting that we want to be applied to the HTML table cells.
        // For example, set a fixed width and a background color.
        builder.CellFormat.Width = 120;                     // Fixed cell width (points)
        builder.CellFormat.Shading.BackgroundPatternColor = System.Drawing.Color.LightYellow;

        // HTML string that contains a simple table.
        const string htmlTable = @"
            <table style='border:1px solid black; border-collapse:collapse;'>
                <tr>
                    <td style='border:1px solid black; padding:5pt;'>Cell 1</td>
                    <td style='border:1px solid black; padding:5pt;'>Cell 2</td>
                </tr>
                <tr>
                    <td style='border:1px solid black; padding:5pt;'>Cell 3</td>
                    <td style='border:1px solid black; padding:5pt;'>Cell 4</td>
                </tr>
            </table>";

        // Insert the HTML table into the document.
        // Use PreserveBlocks to keep block‑level properties (borders, padding, etc.).
        builder.InsertHtml(htmlTable, HtmlInsertOptions.PreserveBlocks);

        // Save the resulting document.
        doc.Save("TableFromHtml.docx");
    }
}
