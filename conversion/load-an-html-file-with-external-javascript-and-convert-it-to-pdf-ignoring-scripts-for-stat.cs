using System;
using System.IO;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Prepare sample HTML with an external JavaScript reference.
        string htmlContent = @"
<!DOCTYPE html>
<html>
<head>
    <title>Sample HTML</title>
    <script src=""script.js""></script>
</head>
<body>
    <h1>Static Heading</h1>
    <p>This paragraph is rendered without executing JavaScript.</p>
</body>
</html>";
        // Write the HTML file.
        const string htmlPath = "input.html";
        File.WriteAllText(htmlPath, htmlContent);

        // Create a dummy external JavaScript file.
        const string jsPath = "script.js";
        File.WriteAllText(jsPath, "console.log('This script is ignored by Aspose.Words.');");

        // Load the HTML document. Aspose.Words ignores JavaScript during rendering.
        Document doc = new Document(htmlPath);

        // Convert and save to PDF.
        const string pdfPath = "output.pdf";
        doc.Save(pdfPath, SaveFormat.Pdf);

        // Validate that the PDF was created.
        if (!File.Exists(pdfPath))
        {
            throw new InvalidOperationException("Expected output PDF was not created.");
        }

        // Optional: Clean up temporary files (comment out if you want to inspect them).
        // File.Delete(htmlPath);
        // File.Delete(jsPath);
    }
}
