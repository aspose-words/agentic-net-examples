using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // Paths for temporary files.
        const string inputPath = "sample.docx";
        const string outputPath = "sample.html";

        // -----------------------------------------------------------------
        // 1. Create a sample DOCX document with some content.
        // -----------------------------------------------------------------
        Document sourceDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(sourceDoc);
        builder.Writeln("This is a sample paragraph.");
        // Add a header to demonstrate round‑trip information.
        builder.MoveToHeaderFooter(HeaderFooterType.HeaderPrimary);
        builder.Writeln("Sample Header");
        // Save the document as DOCX (bootstrap input file).
        sourceDoc.Save(inputPath, SaveFormat.Docx);

        // -----------------------------------------------------------------
        // 2. Load the DOCX document.
        // -----------------------------------------------------------------
        Document doc = new Document(inputPath);

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
        doc.Save(outputPath, htmlOptions);

        // -----------------------------------------------------------------
        // 5. Validate that the HTML file was created.
        // -----------------------------------------------------------------
        if (!File.Exists(outputPath))
            throw new InvalidOperationException("The HTML output file was not created.");

        // Optional: Load the HTML back to verify round‑trip data (e.g., header exists).
        Document roundTripDoc = new Document(outputPath);
        string headerText = roundTripDoc.FirstSection.HeadersFooters[HeaderFooterType.HeaderPrimary].GetText().Trim();
        if (!headerText.Equals("Sample Header", StringComparison.Ordinal))
            throw new InvalidOperationException("Round‑trip information was not preserved.");

        // Clean up temporary files (optional).
        File.Delete(inputPath);
        File.Delete(outputPath);
    }
}
