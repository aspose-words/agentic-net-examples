using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Loading;

class HtmlToPdfConverter
{
    static void Main()
    {
        // Simple HTML content with inline CSS.
        string htmlContent = @"
<!DOCTYPE html>
<html>
<head>
    <style>
        body { font-family: Arial, sans-serif; margin: 40px; }
        h1 { color: #2E86C1; }
        p { font-size: 14pt; }
    </style>
</head>
<body>
    <h1>Sample Document</h1>
    <p>This PDF was generated from HTML while preserving CSS styles.</p>
</body>
</html>";

        // Load the HTML from a memory stream.
        using var htmlStream = new MemoryStream(System.Text.Encoding.UTF8.GetBytes(htmlContent));
        var loadOptions = new LoadOptions { LoadFormat = LoadFormat.Html };
        Document doc = new Document(htmlStream, loadOptions);

        // Determine an output path in the system's temporary folder.
        string pdfPath = Path.Combine(Path.GetTempPath(), "result.pdf");

        // Save the document as PDF.
        doc.Save(pdfPath, SaveFormat.Pdf);

        Console.WriteLine($"PDF successfully created at: {pdfPath}");
    }
}
