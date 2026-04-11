using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

public class ExportRoundTripExample
{
    public static void Main()
    {
        // Define file paths in the current directory.
        string docxPath = Path.Combine(Directory.GetCurrentDirectory(), "Sample.docx");
        string htmlPath = Path.Combine(Directory.GetCurrentDirectory(), "Sample.html");

        // -----------------------------------------------------------------
        // 1. Create a simple DOCX document and save it to disk.
        // -----------------------------------------------------------------
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("Hello world!");
        doc.Save(docxPath, SaveFormat.Docx);

        // -----------------------------------------------------------------
        // 2. Load the saved DOCX document (simulating an existing file).
        // -----------------------------------------------------------------
        Document loadedDoc = new Document(docxPath);

        // -----------------------------------------------------------------
        // 3. Configure HtmlSaveOptions to export round‑trip information.
        // -----------------------------------------------------------------
        HtmlSaveOptions htmlOptions = new HtmlSaveOptions
        {
            ExportRoundtripInformation = true
        };

        // -----------------------------------------------------------------
        // 4. Save the document as HTML using the configured options.
        // -----------------------------------------------------------------
        loadedDoc.Save(htmlPath, htmlOptions);

        // -----------------------------------------------------------------
        // 5. Validate that the HTML file was created and contains round‑trip CSS.
        // -----------------------------------------------------------------
        if (!File.Exists(htmlPath))
            throw new FileNotFoundException("HTML output file was not created.", htmlPath);

        string htmlContent = File.ReadAllText(htmlPath);
        if (string.IsNullOrWhiteSpace(htmlContent))
            throw new InvalidDataException("HTML output file is empty.");

        // The round‑trip information is stored as CSS properties prefixed with "-aw-".
        // Checking for a common marker such as "-aw-headerfooter-type" confirms it.
        if (!htmlContent.Contains("-aw-headerfooter-type"))
            throw new InvalidOperationException("Round‑trip information was not exported to the HTML.");

        // -----------------------------------------------------------------
        // 6. Load the HTML back into a Document to ensure it can be round‑tripped.
        // -----------------------------------------------------------------
        Document roundTripDoc = new Document(htmlPath);

        // Simple verification: the document should contain the same text.
        string roundTripText = roundTripDoc.GetText().Trim();
        if (!roundTripText.Contains("Hello world!"))
            throw new InvalidOperationException("Round‑tripped document does not contain expected text.");

        // If execution reaches this point, the process succeeded.
        Console.WriteLine("ExportRoundTripInformation enabled successfully.");
    }
}
