using System;
using System.IO;
using Aspose.Words;

public class HtmlToPdfConverter
{
    public static void Main()
    {
        // Paths for the temporary HTML input and PDF output files.
        string inputPath = "input.html";
        string outputPath = "output.pdf";

        // Create a simple HTML document that contains embedded CSS styles.
        string htmlContent = @"<!DOCTYPE html>
<html>
<head>
    <meta charset='UTF-8'>
    <title>Sample</title>
    <style>
        .title { color: blue; font-size: 24px; }
        .paragraph { color: green; font-family: Arial; }
    </style>
</head>
<body>
    <h1 class='title'>Hello World</h1>
    <p class='paragraph'>This is a paragraph with CSS styling.</p>
</body>
</html>";

        // Write the HTML string to a local file.
        File.WriteAllText(inputPath, htmlContent);

        // Load the HTML file into an Aspose.Words Document.
        Document doc = new Document(inputPath);

        // Convert the document to PDF while preserving the CSS formatting.
        doc.Save(outputPath, SaveFormat.Pdf);

        // Validate that the PDF file was created successfully.
        if (!File.Exists(outputPath))
        {
            throw new InvalidOperationException("The PDF conversion failed; output file not found.");
        }
    }
}
