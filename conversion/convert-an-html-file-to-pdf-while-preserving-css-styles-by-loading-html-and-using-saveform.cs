using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

public class HtmlToPdfConverter
{
    public static void Main()
    {
        // Define input and output file paths.
        string baseDir = Directory.GetCurrentDirectory();
        string inputHtmlPath = Path.Combine(baseDir, "sample.html");
        string outputPdfPath = Path.Combine(baseDir, "sample.pdf");

        // Create a simple HTML file that contains CSS styles.
        string htmlContent = @"
<!DOCTYPE html>
<html>
<head>
    <meta charset='utf-8'>
    <style>
        body { font-family: Arial, sans-serif; margin: 40px; }
        h1 { color: #2E8B57; }
        p { font-size: 14pt; line-height: 1.5; }
        .highlight { background-color: #FFFF00; }
    </style>
</head>
<body>
    <h1>Aspose.Words HTML to PDF Demo</h1>
    <p>This paragraph demonstrates <span class='highlight'>CSS styling</span> preservation during conversion.</p>
</body>
</html>";

        // Write the HTML content to a local file.
        File.WriteAllText(inputHtmlPath, htmlContent);

        // Load the HTML document.
        Document doc = new Document(inputHtmlPath);

        // Convert and save the document as PDF, preserving CSS styling.
        doc.Save(outputPdfPath, SaveFormat.Pdf);

        // Validate that the PDF file was created.
        if (!File.Exists(outputPdfPath))
        {
            throw new InvalidOperationException($"Failed to create PDF file at '{outputPdfPath}'.");
        }

        // The program finishes here without waiting for user input.
    }
}
