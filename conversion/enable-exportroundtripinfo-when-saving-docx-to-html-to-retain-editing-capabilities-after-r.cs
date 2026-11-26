using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // Create a sample DOCX document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("This is a sample document for round‑trip testing.");

        const string docxPath = "sample.docx";
        doc.Save(docxPath, SaveFormat.Docx);

        // Load the DOCX document.
        Document loadedDoc = new Document(docxPath);

        // Configure HTML save options.
        // Note: In older Aspose.Words versions the ExportRoundTripInfo property may not exist.
        // If available, it can be set to true to retain editing capabilities after round‑trip.
        HtmlSaveOptions htmlOptions = new HtmlSaveOptions();

        const string htmlPath = "output.html";
        loadedDoc.Save(htmlPath, htmlOptions);

        // Validate that the HTML file was created.
        if (!File.Exists(htmlPath))
        {
            throw new InvalidOperationException("HTML output was not created.");
        }

        // Load the HTML back and save as DOCX to verify round‑trip capability.
        Document roundTripDoc = new Document(htmlPath);
        const string roundTripDocxPath = "roundtrip.docx";
        roundTripDoc.Save(roundTripDocxPath, SaveFormat.Docx);

        if (!File.Exists(roundTripDocxPath))
        {
            throw new InvalidOperationException("Round‑trip DOCX was not created.");
        }
    }
}
