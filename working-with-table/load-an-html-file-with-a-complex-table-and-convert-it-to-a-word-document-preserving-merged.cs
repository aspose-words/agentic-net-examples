using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Tables;

public class Program
{
    public static void Main()
    {
        // Prepare directories and file paths.
        string artifactsDir = Path.Combine(Directory.GetCurrentDirectory(), "Artifacts");
        Directory.CreateDirectory(artifactsDir);
        string htmlPath = Path.Combine(artifactsDir, "ComplexTable.html");
        string outputPath = Path.Combine(artifactsDir, "ComplexTable.docx");

        // Create an HTML file that contains a table with merged (colspan/rowspan) cells.
        string html = @"<!DOCTYPE html>
<html>
<head><meta charset='UTF-8'></head>
<body>
<table border='1' cellspacing='0' cellpadding='5'>
  <tr>
    <th colspan='2'>Header spanning two columns</th>
    <th>Header 3</th>
  </tr>
  <tr>
    <td rowspan='2'>Rowspan cell</td>
    <td>Cell 2,1</td>
    <td>Cell 2,2</td>
  </tr>
  <tr>
    <td colspan='2'>Colspan cell</td>
  </tr>
  <tr>
    <td>Cell 4,1</td>
    <td>Cell 4,2</td>
    <td>Cell 4,3</td>
  </tr>
</table>
</body>
</html>";
        File.WriteAllText(htmlPath, html);

        // Load the HTML document. Aspose.Words parses the table and creates merged cells.
        Document doc = new Document(htmlPath);

        // Convert any width‑based merges to explicit merge flags.
        NodeCollection tables = doc.GetChildNodes(NodeType.Table, true);
        foreach (Table table in tables)
        {
            table.ConvertToHorizontallyMergedCells();
        }

        // Save the result as a Word document.
        doc.Save(outputPath, SaveFormat.Docx);

        // Verify that the output file was created.
        if (!File.Exists(outputPath))
            throw new Exception("The Word document was not saved correctly.");
    }
}
