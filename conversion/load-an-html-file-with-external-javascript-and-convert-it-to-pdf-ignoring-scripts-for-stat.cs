using System;
using System.IO;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Create a dummy external JavaScript file.
        const string jsFileName = "script.js";
        File.WriteAllText(jsFileName, "console.log('Hello from external script');");

        // Create an HTML file that references the external JavaScript.
        const string htmlFileName = "input.html";
        string htmlContent = @"
<!DOCTYPE html>
<html>
<head>
    <title>Sample HTML</title>
    <script src='script.js'></script>
</head>
<body>
    <h1>Sample Heading</h1>
    <p>This is a paragraph rendered without executing JavaScript.</p>
</body>
</html>";
        File.WriteAllText(htmlFileName, htmlContent);

        // Load the HTML document. Aspose.Words ignores JavaScript during loading.
        Document doc = new Document(htmlFileName);

        // Convert the loaded document to PDF.
        const string pdfFileName = "output.pdf";
        doc.Save(pdfFileName, SaveFormat.Pdf);

        // Validate that the PDF was created.
        if (!File.Exists(pdfFileName))
            throw new InvalidOperationException("Expected output PDF was not created.");

        // Optional cleanup (comment out if you want to inspect the files after execution).
        // File.Delete(jsFileName);
        // File.Delete(htmlFileName);
        // File.Delete(pdfFileName);
    }
}
