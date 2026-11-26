using System;
using System.IO;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Define paths for the input HTML and output PDF files.
        string htmlPath = "input.html";
        string pdfPath = "output.pdf";

        // Create a simple HTML file with embedded CSS styles.
        string htmlContent = @"
<!DOCTYPE html>
<html>
<head>
    <meta charset='utf-8'>
    <style>
        h1 { color: blue; font-family: Arial; }
        p { font-size: 14pt; font-family: Times New Roman; }
    </style>
</head>
<body>
    <h1>Sample Title</h1>
    <p>This paragraph demonstrates CSS styling in the converted PDF.</p>
</body>
</html>";
        File.WriteAllText(htmlPath, htmlContent);

        // Load the HTML document. Aspose.Words automatically applies the CSS styles.
        Document doc = new Document(htmlPath);

        // Save the loaded document as PDF.
        doc.Save(pdfPath, SaveFormat.Pdf);

        // Validate that the PDF file was created and contains data.
        if (!File.Exists(pdfPath))
            throw new InvalidOperationException("The expected PDF output file was not created.");

        FileInfo pdfInfo = new FileInfo(pdfPath);
        if (pdfInfo.Length == 0)
            throw new InvalidOperationException("The generated PDF file is empty.");

        // Clean up temporary files (optional).
        // File.Delete(htmlPath);
        // File.Delete(pdfPath);
    }
}
