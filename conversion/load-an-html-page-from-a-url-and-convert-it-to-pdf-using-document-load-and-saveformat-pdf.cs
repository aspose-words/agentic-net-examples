using System;
using System.IO;
using Aspose.Words;

public class HtmlToPdfConverter
{
    public static void Main()
    {
        // Create a sample HTML file locally.
        const string htmlFileName = "input.html";
        const string htmlContent = "<html><body><h1>Sample Title</h1><p>This is a sample HTML page.</p></body></html>";
        File.WriteAllText(htmlFileName, htmlContent);

        // Load the HTML document from the local file path.
        string htmlPath = Path.GetFullPath(htmlFileName);
        Document document = new Document(htmlPath);

        // Convert and save to PDF.
        const string pdfFileName = "output.pdf";
        document.Save(pdfFileName, SaveFormat.Pdf);

        // Verify that the PDF file was created.
        if (!File.Exists(pdfFileName))
        {
            throw new InvalidOperationException("Expected output PDF was not created.");
        }
    }
}
