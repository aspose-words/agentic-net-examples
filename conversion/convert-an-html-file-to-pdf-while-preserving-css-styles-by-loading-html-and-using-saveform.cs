using System;
using System.IO;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Define file names.
        const string htmlFile = "input.html";
        const string pdfFile = "output.pdf";

        // Create a simple HTML document with CSS styling.
        string htmlContent = @"
<!DOCTYPE html>
<html>
<head>
    <meta charset=""UTF-8"">
    <title>Sample HTML</title>
    <style>
        body { font-family: Arial, sans-serif; background-color: #f0f0f0; }
        h1 { color: #2e6c80; }
        p { color: #333333; }
        .highlight { background-color: yellow; }
    </style>
</head>
<body>
    <h1>HTML to PDF Conversion</h1>
    <p>This paragraph demonstrates <span class=""highlight"">CSS styling</span> preservation.</p>
</body>
</html>";

        // Write the HTML content to a local file.
        File.WriteAllText(htmlFile, htmlContent);

        // Load the HTML file into an Aspose.Words Document.
        Document doc = new Document(htmlFile);

        // Save the document as PDF, preserving the CSS styles.
        doc.Save(pdfFile, SaveFormat.Pdf);

        // Verify that the PDF was created.
        if (!File.Exists(pdfFile))
            throw new InvalidOperationException("The PDF file was not created.");

        // Optional: Inform the user (no interactive input required).
        Console.WriteLine($"Conversion completed. PDF saved to '{pdfFile}'.");
    }
}
