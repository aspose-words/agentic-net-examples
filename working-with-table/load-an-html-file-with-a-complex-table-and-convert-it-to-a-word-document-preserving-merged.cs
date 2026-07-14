using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Tables;

public class Program
{
    public static void Main()
    {
        // Define paths for the temporary HTML file and the resulting DOCX file.
        string tempFolder = Path.Combine(Directory.GetCurrentDirectory(), "Artifacts");
        Directory.CreateDirectory(tempFolder);

        string htmlPath = Path.Combine(tempFolder, "sample.html");
        string docxPath = Path.Combine(tempFolder, "result.docx");

        // Create an HTML string that contains a complex table with merged cells.
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
            <td colspan='2' style='background:#DDEEFF;'>Horizontally Merged Cell (colspan=2)</td>
            <td style='background:#FFEEDD;'>Cell 3</td>
        </tr>
        <tr>
            <td rowspan='2' style='background:#DDFFDD;'>Vertically Merged Cell (rowspan=2)</td>
            <td style='background:#FFDDDD;'>Cell 2</td>
            <td style='background:#DDDDFF;'>Cell 3</td>
        </tr>
        <tr>
            <td colspan='2' style='background:#FFFFDD;'>Horizontally Merged Cell (colspan=2) in second row</td>
        </tr>
    </table>
</body>
</html>";

        // Write the HTML content to the temporary file.
        File.WriteAllText(htmlPath, htmlContent);

        // Load the HTML file into an Aspose.Words Document.
        Document doc = new Document(htmlPath);

        // Ensure that merged cells are represented using merge flags.
        // This converts cells merged by width (as may happen when loading HTML) to proper CellMerge flags.
        foreach (Table table in doc.GetChildNodes(NodeType.Table, true))
        {
            table.ConvertToHorizontallyMergedCells();
        }

        // Save the document as a DOCX file.
        doc.Save(docxPath, SaveFormat.Docx);

        // Verify that the output file was created.
        if (!File.Exists(docxPath))
        {
            throw new InvalidOperationException("The DOCX file was not created successfully.");
        }

        // Optionally, clean up the temporary HTML file.
        // File.Delete(htmlPath);
    }
}
