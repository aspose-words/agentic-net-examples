using System;
using System.IO;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Prepare a simple HTML content and write it to a local file.
        string htmlContent = "<html><body><h1>Sample</h1><p>Hello world.</p></body></html>";
        string htmlPath = "input.html";
        File.WriteAllText(htmlPath, htmlContent);

        // Load the HTML file into an Aspose.Words Document.
        Document document = new Document(htmlPath);

        // Convert and save the document as PDF.
        string pdfPath = "output.pdf";
        document.Save(pdfPath, SaveFormat.Pdf);

        // Verify that the PDF was created.
        if (!File.Exists(pdfPath))
        {
            throw new InvalidOperationException("Expected output PDF was not created.");
        }
    }
}
