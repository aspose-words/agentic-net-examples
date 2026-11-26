using System;
using System.IO;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Create a simple external JavaScript file.
        const string jsFileName = "script.js";
        File.WriteAllText(jsFileName, "console.log('Hello from script');");

        // Create an HTML file that references the external JavaScript.
        const string htmlFileName = "input.html";
        string htmlContent = @"<html>
<head>
    <script src=""script.js""></script>
</head>
<body>
    <h1>Sample</h1>
    <p>Hello world.</p>
</body>
</html>";
        File.WriteAllText(htmlFileName, htmlContent);

        // Load the HTML document. Aspose.Words does not execute JavaScript,
        // so the script is effectively ignored during conversion.
        Document doc = new Document(htmlFileName);

        // Convert the loaded document to PDF.
        const string pdfFileName = "output.pdf";
        doc.Save(pdfFileName, SaveFormat.Pdf);

        // Validate that the PDF was created.
        if (!File.Exists(pdfFileName))
            throw new InvalidOperationException("Expected output PDF was not created.");

        // Optionally, clean up the temporary files (comment out if inspection is needed).
        // File.Delete(jsFileName);
        // File.Delete(htmlFileName);
    }
}
