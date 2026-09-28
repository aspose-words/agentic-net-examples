using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Choose a font that does not contain OpenType features (e.g., Arial).
        // This effectively disables ligatures and other OpenType behaviors for the rendered text.
        builder.Font.Name = "Arial";
        builder.Font.Size = 24;

        // Add text that would normally display ligatures if OpenType features were enabled.
        builder.Writeln("Office");
        builder.Writeln("affinity");
        builder.Writeln("fluff");

        // Render the document to PDF.
        string pdfPath = "output.pdf";
        doc.Save(pdfPath, SaveFormat.Pdf);

        // Verify that the PDF file was created successfully.
        if (!File.Exists(pdfPath))
        {
            throw new InvalidOperationException("Failed to create the PDF file.");
        }

        // Indicate successful completion (no user interaction required).
        Console.WriteLine($"PDF generated at: {Path.GetFullPath(pdfPath)}");
    }
}
