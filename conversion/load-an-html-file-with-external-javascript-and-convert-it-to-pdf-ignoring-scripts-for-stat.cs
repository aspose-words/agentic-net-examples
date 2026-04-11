using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Loading;
using Aspose.Words.Saving;

public class HtmlToPdfExample
{
    public static void Main()
    {
        // Prepare a temporary folder for the sample files.
        string baseDir = Path.Combine(Directory.GetCurrentDirectory(), "SampleData");
        Directory.CreateDirectory(baseDir);

        // Create an external JavaScript file.
        string scriptPath = Path.Combine(baseDir, "script.js");
        File.WriteAllText(scriptPath, "function showMessage() { alert('Hello from script!'); }");

        // Create an HTML file that references the external script.
        string htmlContent = $@"
<!DOCTYPE html>
<html>
<head>
    <title>Sample HTML</title>
    <script src=""{Path.GetFileName(scriptPath)}""></script>
</head>
<body>
    <h1>Static Content</h1>
    <p>This paragraph should appear in the PDF. The script is ignored.</p>
</body>
</html>";
        string htmlPath = Path.Combine(baseDir, "sample.html");
        File.WriteAllText(htmlPath, htmlContent);

        // Load the HTML document. Aspose.Words ignores script execution by default.
        using (FileStream htmlStream = new FileStream(htmlPath, FileMode.Open, FileAccess.Read))
        {
            HtmlLoadOptions loadOptions = new HtmlLoadOptions();
            // No special option is required to ignore scripts; they are not processed.
            Document doc = new Document(htmlStream, loadOptions);

            // Define the output PDF path.
            string pdfPath = Path.Combine(baseDir, "output.pdf");

            // Save the document as PDF.
            doc.Save(pdfPath, SaveFormat.Pdf);

            // Validate that the PDF was created and contains data.
            if (!File.Exists(pdfPath))
                throw new FileNotFoundException("PDF file was not created.", pdfPath);

            FileInfo pdfInfo = new FileInfo(pdfPath);
            if (pdfInfo.Length == 0)
                throw new InvalidOperationException("Generated PDF file is empty.");
        }

        // Clean up (optional): uncomment the following line to delete the sample files after execution.
        // Directory.Delete(baseDir, true);
    }
}
