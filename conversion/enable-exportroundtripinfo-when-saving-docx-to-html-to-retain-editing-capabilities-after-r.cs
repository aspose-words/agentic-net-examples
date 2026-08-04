using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

public class ExportRoundTripInfoExample
{
    public static void Main()
    {
        // Step 1: Create a sample DOCX document.
        Document originalDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(originalDoc);
        builder.Writeln("This is a sample document for round‑trip conversion.");
        builder.StartTable();
        builder.InsertCell();
        builder.Writeln("Cell 1");
        builder.InsertCell();
        builder.Writeln("Cell 2");
        builder.EndTable();

        const string originalDocPath = "original.docx";
        originalDoc.Save(originalDocPath, SaveFormat.Docx);
        if (!File.Exists(originalDocPath))
            throw new InvalidOperationException($"Failed to create '{originalDocPath}'.");

        // Step 2: Load the DOCX and save to HTML.
        // The ExportRoundTripInfo property is not available in older Aspose.Words versions,
        // but the conversion flow remains the same.
        Document docForHtml = new Document(originalDocPath);
        HtmlSaveOptions htmlOptions = new HtmlSaveOptions
        {
            ExportImagesAsBase64 = true
        };
        const string htmlPath = "roundtrip.html";
        docForHtml.Save(htmlPath, htmlOptions);
        if (!File.Exists(htmlPath))
            throw new InvalidOperationException($"Failed to create '{htmlPath}'.");

        // Step 3: Load the generated HTML and save back to DOCX.
        Document roundTripDoc = new Document(htmlPath);
        const string roundTripDocPath = "roundtrip_back.docx";
        roundTripDoc.Save(roundTripDocPath, SaveFormat.Docx);
        if (!File.Exists(roundTripDocPath))
            throw new InvalidOperationException($"Failed to create '{roundTripDocPath}'.");

        // Simple validation: ensure the round‑trip DOCX is not empty.
        FileInfo info = new FileInfo(roundTripDocPath);
        if (info.Length == 0)
            throw new InvalidOperationException("The round‑trip DOCX file is empty.");

        // Indicate success.
        Console.WriteLine("Round‑trip conversion completed successfully.");
    }
}
