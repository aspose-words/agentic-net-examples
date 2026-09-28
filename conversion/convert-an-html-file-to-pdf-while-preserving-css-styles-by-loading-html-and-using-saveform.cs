using System;
using System.IO;
using Aspose.Words;

public class HtmlToPdfConverter
{
    public static void Main()
    {
        // Create a sample HTML file with CSS styles.
        const string htmlPath = "input.html";
        const string htmlContent = @"
<!DOCTYPE html>
<html>
<head>
    <style>
        body { font-family: Arial, sans-serif; margin: 40px; }
        h1 { color: #2E86C1; }
        p { font-size: 14pt; line-height: 1.5; }
        .highlight { background-color: #FFF9C4; padding: 5px; }
    </style>
</head>
<body>
    <h1>Sample Document</h1>
    <p>This paragraph demonstrates <span class=""highlight"">styled text</span> conversion from HTML to PDF using Aspose.Words.</p>
</body>
</html>";
        File.WriteAllText(htmlPath, htmlContent);

        // Load the HTML file into an Aspose.Words Document.
        Document document = new Document(htmlPath);

        // Save the document as PDF, preserving the CSS styles.
        const string pdfPath = "output.pdf";
        document.Save(pdfPath, SaveFormat.Pdf);

        // Validate that the PDF file was created.
        if (!File.Exists(pdfPath))
        {
            throw new InvalidOperationException("The PDF file was not created as expected.");
        }

        // Optionally, clean up the temporary HTML file.
        // File.Delete(htmlPath);
    }
}
