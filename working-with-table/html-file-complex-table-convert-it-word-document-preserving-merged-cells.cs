using System;
using System.IO;
using System.Text;
using Aspose.Words;
using Aspose.Words.Tables;
using Aspose.Words.Loading;

class HtmlToWordConverter
{
    static void Main()
    {
        // Sample HTML containing a complex table with merged cells.
        const string htmlContent = @"
<!DOCTYPE html>
<html>
<head>
    <meta charset='utf-8' />
    <title>Complex Table</title>
    <style>
        table { border-collapse: collapse; width: 100%; }
        td, th { border: 1px solid #000; padding: 5px; }
    </style>
</head>
<body>
    <table>
        <tr>
            <th colspan='2'>Header spanning two columns</th>
        </tr>
        <tr>
            <td rowspan='2'>Cell merged vertically</td>
            <td>Cell 1</td>
        </tr>
        <tr>
            <td>Cell 2</td>
        </tr>
        <tr>
            <td colspan='2'>Cell merged horizontally</td>
        </tr>
    </table>
</body>
</html>";

        // Load the HTML document from a memory stream.
        var loadOptions = new LoadOptions { LoadFormat = LoadFormat.Html };
        using var stream = new MemoryStream(Encoding.UTF8.GetBytes(htmlContent));
        Document doc = new Document(stream, loadOptions);

        // Convert cells merged by width into horizontally merged cells to preserve layout.
        foreach (Table table in doc.GetChildNodes(NodeType.Table, true))
        {
            table.ConvertToHorizontallyMergedCells();
        }

        // Save the resulting document in the current working directory.
        string outputFilePath = Path.Combine(Environment.CurrentDirectory, "ConvertedDocument.docx");
        doc.Save(outputFilePath, SaveFormat.Docx);

        Console.WriteLine($"Document saved to: {outputFilePath}");
    }
}
