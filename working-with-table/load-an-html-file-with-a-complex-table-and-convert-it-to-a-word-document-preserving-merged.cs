using System;
using System.IO;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Create a temporary folder for the sample files.
        string tempFolder = Path.Combine(Path.GetTempPath(), "AsposeWordsSample_" + Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(tempFolder);

        // Define the HTML content with a complex table (merged cells).
        string htmlContent = @"
<!DOCTYPE html>
<html>
<head>
    <meta charset='UTF-8'>
    <title>Sample Table</title>
</head>
<body>
    <table border='1' style='border-collapse:collapse;'>
        <tr>
            <th colspan='2'>Header 1-2</th>
            <th>Header 3</th>
        </tr>
        <tr>
            <td rowspan='2'>Rowspan Cell</td>
            <td>Cell 2,1</td>
            <td>Cell 2,2</td>
        </tr>
        <tr>
            <td colspan='2'>Colspan Cell</td>
        </tr>
    </table>
</body>
</html>";

        // Write the HTML to a file.
        string htmlPath = Path.Combine(tempFolder, "sample.html");
        File.WriteAllText(htmlPath, htmlContent);

        // Load the HTML file into an Aspose.Words Document.
        Document doc = new Document(htmlPath);

        // Save the document as a Word file.
        string outputPath = Path.Combine(tempFolder, "output.docx");
        doc.Save(outputPath);

        // Validate that the output file was created.
        if (!File.Exists(outputPath))
        {
            throw new InvalidOperationException("The Word document was not created as expected.");
        }

        // Clean up temporary files (optional).
        // Comment out the following lines if you want to inspect the files after execution.
        try
        {
            File.Delete(htmlPath);
            File.Delete(outputPath);
            Directory.Delete(tempFolder);
        }
        catch
        {
            // Ignored – cleanup failures should not affect program outcome.
        }
    }
}
