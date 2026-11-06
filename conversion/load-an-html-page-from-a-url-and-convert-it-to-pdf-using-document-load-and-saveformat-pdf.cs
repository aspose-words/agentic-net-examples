using System;
using System.IO;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Create a simple HTML file locally.
        const string htmlFileName = "input.html";
        string htmlContent = "<html><body><h1>Sample HTML</h1><p>This is a test.</p></body></html>";
        File.WriteAllText(htmlFileName, htmlContent);

        // Load the HTML document from the local file.
        Document document = new Document(htmlFileName);

        // Convert and save the document as PDF.
        const string pdfFileName = "output.pdf";
        document.Save(pdfFileName, SaveFormat.Pdf);

        // Validate that the PDF was created.
        if (!File.Exists(pdfFileName))
        {
            throw new InvalidOperationException("Expected output PDF was not created.");
        }
    }
}
