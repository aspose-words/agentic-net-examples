using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Loading;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // Define file paths in the current working directory.
        string baseDir = Directory.GetCurrentDirectory();
        string htmlPath = Path.Combine(baseDir, "input.html");
        string jsPath = Path.Combine(baseDir, "script.js");
        string pdfPath = Path.Combine(baseDir, "output.pdf");

        // Create a simple external JavaScript file.
        string jsContent = "function greet() { console.log('Hello from external script'); }";
        File.WriteAllText(jsPath, jsContent);

        // Create an HTML file that references the external JavaScript.
        string htmlContent = $@"<!DOCTYPE html>
<html>
<head>
    <title>Sample HTML</title>
    <script src=""{Path.GetFileName(jsPath)}""></script>
</head>
<body>
    <h1>Hello World</h1>
    <p>This document is converted to PDF while ignoring scripts.</p>
</body>
</html>";
        File.WriteAllText(htmlPath, htmlContent);

        // Load the HTML document. Scripts are not executed by Aspose.Words,
        // but we set IgnoreNoscriptElements to true to demonstrate ignoring
        // any <noscript> content that might be present.
        HtmlLoadOptions loadOptions = new HtmlLoadOptions
        {
            IgnoreNoscriptElements = true
        };
        Document doc = new Document(htmlPath, loadOptions);

        // Convert the loaded document to PDF.
        doc.Save(pdfPath, SaveFormat.Pdf);

        // Verify that the PDF file was created.
        if (!File.Exists(pdfPath))
        {
            throw new InvalidOperationException("The PDF output file was not created.");
        }
    }
}
