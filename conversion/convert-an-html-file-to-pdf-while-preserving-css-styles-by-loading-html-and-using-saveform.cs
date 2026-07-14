using System;
using System.IO;
using Aspose.Words;

public class HtmlToPdfConverter
{
    public static void Main()
    {
        // Define file names.
        const string htmlPath = "sample.html";
        const string pdfPath = "sample.pdf";

        // Create a simple HTML file with CSS styling.
        string htmlContent = @"
<!DOCTYPE html>
<html>
<head>
    <meta charset='UTF-8'>
    <title>Sample HTML</title>
    <style>
        body { font-family: Arial, sans-serif; margin: 40px; }
        h1 { color: #2E86C1; }
        p { color: #555555; line-height: 1.5; }
        .highlight { background-color: #FFF9C4; }
    </style>
</head>
<body>
    <h1>HTML to PDF Conversion</h1>
    <p>This paragraph demonstrates <span class='highlight'>CSS styling</span> preservation during conversion.</p>
</body>
</html>";

        // Write the HTML content to a local file.
        File.WriteAllText(htmlPath, htmlContent);

        // Load the HTML document.
        Document doc = new Document(htmlPath);

        // Save the document as PDF, preserving CSS styles.
        doc.Save(pdfPath, SaveFormat.Pdf);

        // Verify that the PDF file was created.
        if (!File.Exists(pdfPath))
            throw new InvalidOperationException("The PDF file was not created.");

        // Optional: clean up the temporary HTML file.
        // File.Delete(htmlPath);
    }
}
