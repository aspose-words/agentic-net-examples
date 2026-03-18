using System;
using Aspose.Words;
using Aspose.Words.Saving;

public class DocxToHtmlConverter
{
    /// <summary>
    /// Converts a DOCX file that contains complex tables to HTML.
    /// The conversion preserves table borders, cell padding and other layout details.
    /// </summary>
    /// <param name="inputDocxPath">Full path to the source DOCX file.</param>
    /// <param name="outputHtmlPath">Full path where the resulting HTML file will be saved.</param>
    public static void ConvertComplexTablesToHtml(string inputDocxPath, string outputHtmlPath)
    {
        // Load the DOCX document from the file system.
        Document doc = new Document(inputDocxPath);

        // Expand any table style formatting to direct formatting.
        // This ensures that style‑based borders and padding are written explicitly to the HTML.
        doc.ExpandTableStylesToDirectFormatting();

        // Configure HTML save options.
        HtmlSaveOptions saveOptions = new HtmlSaveOptions(SaveFormat.Html)
        {
            // Preserve round‑trip information (borders, padding, hidden elements, etc.).
            ExportRoundtripInformation = true,

            // Export all width values (absolute and relative) so that table layout matches the source.
            TableWidthOutputMode = HtmlElementSizeOutputMode.All,

            // Write the DOCTYPE declaration for XHTML Transitional (optional but keeps full markup).
            ExportXhtmlTransitional = true,

            // Keep negative indents if any tables use them.
            AllowNegativeIndent = true
        };

        // Save the document as HTML using the configured options.
        doc.Save(outputHtmlPath, saveOptions);
    }

    // ---------------------------------------------------------------------
    // Entry point required for a console application.
    // ---------------------------------------------------------------------
    public static void Main(string[] args)
    {
        // Simple argument handling – if none are supplied, show usage.
        if (args.Length != 2)
        {
            Console.WriteLine("Usage: DocxToHtmlConverter <input.docx> <output.html>");
            return;
        }

        string inputPath = args[0];
        string outputPath = args[1];

        try
        {
            ConvertComplexTablesToHtml(inputPath, outputPath);
            Console.WriteLine($"Conversion succeeded. HTML saved to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error during conversion: {ex.Message}");
        }
    }
}
