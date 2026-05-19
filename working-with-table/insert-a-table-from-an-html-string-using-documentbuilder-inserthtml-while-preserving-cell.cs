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
            // Create a new empty document.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // HTML string that contains a table with cell background colors and widths.
            const string html = @"
                <table style='border:1px solid black; border-collapse:collapse;'>
                    <tr>
                        <td style='background:#FFCCCC; width:100pt; padding:5pt;'>Cell 1</td>
                        <td style='background:#CCFFCC; width:150pt; padding:5pt;'>Cell 2</td>
                    </tr>
                </table>";

            // Insert the HTML fragment into the document.
            builder.InsertHtml(html);

            // Retrieve the first table that was inserted.
            Table table = (Table)doc.GetChild(NodeType.Table, 0, true);
            if (table == null)
                throw new InvalidOperationException("No table was inserted from the HTML.");

            // Validate that the cell formatting (background color) was preserved.
            Color expectedFirstCellColor = ColorTranslator.FromHtml("#FFCCCC");
            Color expectedSecondCellColor = ColorTranslator.FromHtml("#CCFFCC");

            Cell firstCell = table.FirstRow.FirstCell;
            Cell secondCell = table.FirstRow.LastCell;

            if (firstCell.CellFormat.Shading.BackgroundPatternColor.ToArgb() != expectedFirstCellColor.ToArgb())
                throw new InvalidOperationException("First cell background color does not match expected value.");

            if (secondCell.CellFormat.Shading.BackgroundPatternColor.ToArgb() != expectedSecondCellColor.ToArgb())
                throw new InvalidOperationException("Second cell background color does not match expected value.");

            // Save the document to a file.
            string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "OutputTableFromHtml.docx");
            doc.Save(outputPath);

            // Verify that the file was created.
            if (!File.Exists(outputPath))
                throw new FileNotFoundException("The output document was not saved.", outputPath);
        }
    }
}
