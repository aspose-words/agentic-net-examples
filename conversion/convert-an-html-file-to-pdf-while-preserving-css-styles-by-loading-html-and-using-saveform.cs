using System;
using System.IO;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Create a sample HTML file that includes CSS styling.
        const string htmlPath = "input.html";
        const string pdfPath = "output.pdf";
        string htmlContent = @"
<!DOCTYPE html>
<html>
<head>
    <style>
        body { font-family: Arial; color: #333333; }
        h1 { color: #0066CC; }
    </style>
</head>
<body>
    <h1>Hello World</h1>
    <p>This is a sample HTML document with CSS styles.</p>
</body>
</html>";
        File.WriteAllText(htmlPath, htmlContent);

        // Load the HTML file into an Aspose.Words Document.
        Document doc = new Document(htmlPath);

        // Convert the document to PDF while preserving the CSS styles.
        doc.Save(pdfPath, SaveFormat.Pdf);

        // Verify that the PDF file was created.
        if (!File.Exists(pdfPath))
            throw new InvalidOperationException("Expected output PDF was not created.");
    }
}
