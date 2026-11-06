using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Fonts;   // Required for FontSettings
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Add a paragraph containing characters that normally form OpenType ligatures.
        builder.Writeln("Office and file have ligatures like fi and fl.");

        // Configure FontSettings (no OpenType disabling API is available in the current version).
        FontSettings fontSettings = new FontSettings();
        doc.FontSettings = fontSettings;

        // Define the output PDF path.
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "Output.pdf");

        // Render the document to PDF.
        PdfSaveOptions saveOptions = new PdfSaveOptions();
        doc.Save(outputPath, saveOptions);

        // Verify that the PDF file was created.
        if (!File.Exists(outputPath))
            throw new FileNotFoundException("The PDF file was not created.", outputPath);

        // Output the file size for confirmation.
        long fileSize = new FileInfo(outputPath).Length;
        Console.WriteLine($"PDF generated successfully. File size: {fileSize} bytes.");
    }
}
