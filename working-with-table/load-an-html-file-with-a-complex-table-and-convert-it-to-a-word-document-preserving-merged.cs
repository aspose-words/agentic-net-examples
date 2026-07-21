using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Tables;

public class HtmlToWordConverter
{
    public static void Main()
    {
        // Define HTML content with a complex table that includes merged cells.
        string htmlContent = @"
<!DOCTYPE html>
<html>
<head>
    <meta charset='UTF-8'>
    <title>Sample Table</title>
</head>
<body>
    <table border='1' cellspacing='0' cellpadding='5'>
        <tr>
            <th rowspan='2'>Header 1</th>
            <th colspan='2'>Header 2-3</th>
            <th>Header 4</th>
        </tr>
        <tr>
            <th>Subheader 2</th>
            <th>Subheader 3</th>
            <th>Subheader 4</th>
        </tr>
        <tr>
            <td>R1C1</td>
            <td colspan='2'>R1C2-3 merged</td>
            <td>R1C4</td>
        </tr>
        <tr>
            <td rowspan='2'>R2-3C1 merged</td>
            <td>R2C2</td>
            <td>R2C3</td>
            <td>R2C4</td>
        </tr>
        <tr>
            <td colspan='2'>R3C2-3 merged</td>
            <td>R3C4</td>
        </tr>
    </table>
</body>
</html>";

        // Create a temporary HTML file.
        string tempHtmlPath = Path.Combine(Path.GetTempPath(), "sample_table.html");
        File.WriteAllText(tempHtmlPath, htmlContent);

        // Load the HTML file into an Aspose.Words Document.
        Document doc = new Document(tempHtmlPath);

        // Ensure that the document contains at least one table.
        if (doc.FirstSection?.Body?.Tables?.Count == 0)
            throw new InvalidOperationException("No tables were found in the loaded HTML document.");

        // Convert any width‑based merged cells to proper merge flags.
        Table table = doc.FirstSection.Body.Tables[0];
        table.ConvertToHorizontallyMergedCells();

        // Save the resulting document as a Word file.
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "ConvertedTable.docx");
        doc.Save(outputPath, SaveFormat.Docx);

        // Verify that the output file was created.
        if (!File.Exists(outputPath))
            throw new IOException($"Failed to create the output file at '{outputPath}'.");

        // Clean up the temporary HTML file.
        File.Delete(tempHtmlPath);
    }
}
