using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Loading;

public class Program
{
    public static void Main()
    {
        // Define file names
        const string htmlFile = "sample.html";
        const string jsFile = "script.js";
        const string pdfFile = "output.pdf";

        // Create a simple external JavaScript file
        File.WriteAllText(jsFile, @"
function showMessage() {
    // This script will not affect the static PDF rendering
    console.log('Hello from external script!');
}
");

        // Create an HTML file that references the external JavaScript
        string htmlContent = $@"
<!DOCTYPE html>
<html>
<head>
    <title>Sample HTML with External JS</title>
    <script src=""{jsFile}""></script>
</head>
<body>
    <h1>Static Content</h1>
    <p>This paragraph will appear in the PDF. The script is ignored during conversion.</p>
</body>
</html>";
        File.WriteAllText(htmlFile, htmlContent);

        // Load the HTML document, ignoring <noscript> elements (scripts are ignored by default)
        HtmlLoadOptions loadOptions = new HtmlLoadOptions
        {
            IgnoreNoscriptElements = true
        };
        Document doc = new Document(htmlFile, loadOptions);

        // Convert to PDF
        doc.Save(pdfFile, SaveFormat.Pdf);

        // Validate that the PDF was created
        if (!File.Exists(pdfFile) || new FileInfo(pdfFile).Length == 0)
            throw new InvalidOperationException("PDF conversion failed: output file was not created or is empty.");

        // Clean up temporary files (optional)
        // File.Delete(htmlFile);
        // File.Delete(jsFile);
    }
}
