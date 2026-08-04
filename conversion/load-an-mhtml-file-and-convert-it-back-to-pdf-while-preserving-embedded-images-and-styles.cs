using System;
using System.IO;
using System.Text;
using Aspose.Words;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // Define file names.
        const string htmlFile = "sample.html";
        const string mhtmlFile = "sample.mhtml";
        const string pdfFile = "output.pdf";

        // Create a simple HTML document with inline CSS and an embedded image (base64 PNG).
        string htmlContent =
            "<html>" +
            "<head>" +
            "<style>h1 { color: blue; }</style>" +
            "</head>" +
            "<body>" +
            "<h1>Hello, Aspose.Words!</h1>" +
            // 1x1 pixel transparent PNG.
            "<img src=\"data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAwMCAO+XK6cAAAAASUVORK5CYII=\" alt=\"pixel\"/>" +
            "</body>" +
            "</html>";

        // Write the HTML to a temporary file (optional, but keeps the example clear).
        File.WriteAllText(htmlFile, htmlContent, Encoding.UTF8);

        // Load the HTML into an Aspose.Words Document.
        Document docFromHtml = new Document(htmlFile);

        // Save the document as MHTML, preserving the embedded image and styles.
        docFromHtml.Save(mhtmlFile, SaveFormat.Mhtml);

        // Verify that the MHTML file was created.
        if (!File.Exists(mhtmlFile))
            throw new InvalidOperationException("MHTML file was not created.");

        // Load the MHTML file.
        Document docFromMhtml = new Document(mhtmlFile);

        // Convert the loaded document to PDF, preserving images and styles.
        docFromMhtml.Save(pdfFile, SaveFormat.Pdf);

        // Verify that the PDF file was created.
        if (!File.Exists(pdfFile))
            throw new InvalidOperationException("PDF file was not created.");

        // Clean up temporary files (optional).
        File.Delete(htmlFile);
        File.Delete(mhtmlFile);
    }
}
