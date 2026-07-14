using System;
using System.IO;
using System.Drawing;
using Aspose.Words;
using Aspose.Words.Tables;

namespace AsposeWordsTableFromHtml
{
    public class Program
    {
        public static void Main()
        {
            // Create a new blank document.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // HTML string that contains a table with cell background colors and explicit widths.
            const string html = @"
                <table style='border-collapse:collapse;'>
                    <tr>
                        <td style='background:#FFCCCC; width:100pt; height:30pt;'>Cell 1</td>
                        <td style='background:#CCFFCC; width:100pt; height:30pt;'>Cell 2</td>
                    </tr>
                </table>";

            // Insert the HTML fragment into the document. The builder parses the HTML and creates a table.
            builder.InsertHtml(html);

            // Define the output path in the current working directory.
            string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "TableFromHtml.docx");

            // Save the document.
            doc.Save(outputPath);

            // Verify that the file was created.
            if (!File.Exists(outputPath))
                throw new Exception("The output document was not saved.");

            // Retrieve the first table that was inserted from the HTML.
            Table table = (Table)doc.GetChildNodes(NodeType.Table, true)[0];

            // Verify that the first cell retains the background color defined in the HTML.
            Cell firstCell = table.FirstRow.FirstCell;
            Color expectedColor = Color.FromArgb(255, 255, 204, 204); // #FFCCCC
            Color actualColor = firstCell.CellFormat.Shading.BackgroundPatternColor;

            if (actualColor.ToArgb() != expectedColor.ToArgb())
                throw new Exception($"Cell background color mismatch. Expected: {expectedColor}, Actual: {actualColor}");

            // Verify that the second cell also retains its background color.
            Cell secondCell = table.FirstRow.LastCell;
            Color expectedSecondColor = Color.FromArgb(255, 204, 255, 204); // #CCFFCC
            Color actualSecondColor = secondCell.CellFormat.Shading.BackgroundPatternColor;

            if (actualSecondColor.ToArgb() != expectedSecondColor.ToArgb())
                throw new Exception($"Second cell background color mismatch. Expected: {expectedSecondColor}, Actual: {actualSecondColor}");

            // Program ends without further interaction.
        }
    }
}
