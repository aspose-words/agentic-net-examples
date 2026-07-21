using System;
using System.IO;
using Aspose.Words;

public class HtmlToPdfConverter
{
    public static void Main()
    {
        // Paths for the temporary HTML input and PDF output.
        const string inputHtmlPath = "input.html";
        const string outputPdfPath = "output.pdf";

        // Sample HTML that includes CSS styling.
        string htmlContent = @"<!DOCTYPE html>
<html>
<head>
<style>
h1 { color: blue; }
p { font-size: 14pt; }
</style>
</head>
<body>
<h1>Sample Heading</h1>
<p>This is a paragraph with CSS styling.</p>
</body>
</html>";

        // Write the HTML content to a local file.
        File.WriteAllText(inputHtmlPath, htmlContent);

        // Load the HTML file into an Aspose.Words Document.
        Document document = new Document(inputHtmlPath);

        // Convert and save the document as PDF, preserving the CSS styles.
        document.Save(outputPdfPath, SaveFormat.Pdf);

        // Validate that the PDF file was created.
        if (!File.Exists(outputPdfPath))
        {
            throw new InvalidOperationException("The PDF output file was not created.");
        }
    }
}
