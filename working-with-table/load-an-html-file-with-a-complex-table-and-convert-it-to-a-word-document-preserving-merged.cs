using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Tables;

public class Program
{
    public static void Main()
    {
        // Prepare output directory.
        string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "Output");
        Directory.CreateDirectory(outputDir);

        // Create a sample HTML file containing a complex table with merged cells.
        string htmlPath = Path.Combine(outputDir, "SampleTable.html");
        string htmlContent = @"
<!DOCTYPE html>
<html>
<head>
    <meta charset='UTF-8'>
    <title>Complex Table</title>
</head>
<body>
    <table border='1' cellspacing='0' cellpadding='5'>
        <tr>
            <th colspan='2'>Header Merged Horizontally</th>
            <th>Header 3</th>
        </tr>
        <tr>
            <td rowspan='2'>Cell Merged Vertically</td>
            <td>Row 1, Cell 2</td>
            <td>Row 1, Cell 3</td>
        </tr>
        <tr>
            <td colspan='2'>Cell Merged Horizontally</td>
        </tr>
    </table>
</body>
</html>";
        File.WriteAllText(htmlPath, htmlContent);

        // Load the HTML file into an Aspose.Words Document.
        Document doc = new Document(htmlPath);

        // Convert any cells that were merged by width into proper merge flags.
        NodeCollection tables = doc.GetChildNodes(NodeType.Table, true);
        foreach (Table table in tables)
        {
            table.ConvertToHorizontallyMergedCells();
        }

        // Save the document as a Word file.
        string outputPath = Path.Combine(outputDir, "Converted.docx");
        doc.Save(outputPath, SaveFormat.Docx);

        // Verify that the output file was created.
        if (!File.Exists(outputPath))
            throw new Exception("The Word document was not saved successfully.");
    }
}
