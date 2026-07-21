using System;
using System.Drawing;
using System.IO;
using Aspose.Words;
using Aspose.Words.Tables;
using Aspose.Words.Saving;

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
                    <td style='background:#FFCCCC; width:100pt;'>Cell 1</td>
                    <td style='background:#CCFFCC; width:100pt;'>Cell 2</td>
                </tr>
            </table>";

        // Insert the HTML into the document while preserving block‑level formatting (e.g., cell shading).
        builder.InsertHtml(html, HtmlInsertOptions.PreserveBlocks);

        // Save the document to a local file.
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "InsertedTable.docx");
        doc.Save(outputPath);

        // Verify that the file was created.
        if (!File.Exists(outputPath))
            throw new InvalidOperationException("The output document was not saved.");

        // Retrieve the inserted table.
        Table table = doc.FirstSection.Body.Tables[0];
        if (table == null)
            throw new InvalidOperationException("No table was found after HTML insertion.");

        // Verify that the first cell retained its background color from the HTML.
        Cell firstCell = table.Rows[0].Cells[0];
        Color expectedColor = ColorTranslator.FromHtml("#FFCCCC");
        Color actualColor = firstCell.CellFormat.Shading.BackgroundPatternColor;

        if (actualColor.ToArgb() != expectedColor.ToArgb())
            throw new InvalidOperationException("Cell background color was not preserved.");

        // Inform that the process completed successfully.
        Console.WriteLine("Document created and table formatting verified successfully.");
    }
}
