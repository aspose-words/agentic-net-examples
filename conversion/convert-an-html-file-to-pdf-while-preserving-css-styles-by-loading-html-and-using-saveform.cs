using System;
using System.IO;
using Aspose.Words;

public class HtmlToPdfConverter
{
    public static void Main()
    {
        // Prepare a sample HTML file with CSS styling.
        const string htmlFileName = "input.html";
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
    <h1>Sample Heading</h1>
    <p>This is a paragraph with <span class='highlight'>highlighted text</span> to demonstrate CSS styling preservation.</p>
</body>
</html>";
        File.WriteAllText(htmlFileName, htmlContent);

        // Load the HTML document.
        Document document = new Document(htmlFileName);

        // Convert and save the document as PDF.
        const string pdfFileName = "output.pdf";
        document.Save(pdfFileName, SaveFormat.Pdf);

        // Validate that the PDF was created.
        if (!File.Exists(pdfFileName))
        {
            throw new InvalidOperationException("Expected output PDF was not created.");
        }

        // Optional: clean up the temporary HTML file.
        // File.Delete(htmlFileName);
    }
}
