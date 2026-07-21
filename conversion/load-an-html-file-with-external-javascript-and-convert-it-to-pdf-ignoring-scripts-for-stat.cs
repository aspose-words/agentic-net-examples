using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Loading;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // Define file names.
        string htmlFile = "input.html";
        string jsFile = "script.js";
        string pdfFile = "output.pdf";

        // Create a simple external JavaScript file.
        File.WriteAllText(jsFile, "function hello() { alert('Hello from external script!'); }");

        // Create an HTML file that references the external JavaScript.
        string htmlContent = @"<!DOCTYPE html>
<html>
<head>
    <title>Sample HTML</title>
    <script src=""script.js""></script>
</head>
<body>
    <h1>Hello World</h1>
    <p>This document is converted to PDF while ignoring scripts.</p>
</body>
</html>";
        File.WriteAllText(htmlFile, htmlContent);

        // Load the HTML document. Scripts are not executed by Aspose.Words.
        HtmlLoadOptions loadOptions = new HtmlLoadOptions
        {
            // BaseUri helps resolve relative paths (e.g., the script file).
            BaseUri = Directory.GetCurrentDirectory(),
            // Optional: ignore <noscript> elements if present.
            IgnoreNoscriptElements = true
        };

        Document doc = new Document(htmlFile, loadOptions);

        // Convert the loaded document to PDF.
        doc.Save(pdfFile, SaveFormat.Pdf);

        // Verify that the PDF was created.
        if (!File.Exists(pdfFile))
            throw new InvalidOperationException("The PDF file was not created.");

        // Clean up temporary files (optional).
        File.Delete(htmlFile);
        File.Delete(jsFile);
    }
}
