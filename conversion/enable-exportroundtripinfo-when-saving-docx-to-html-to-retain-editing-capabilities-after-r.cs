using System;
using System.IO;
using System.Reflection;
using Aspose.Words;
using Aspose.Words.Saving;

public class ExportRoundTripInfoExample
{
    public static void Main()
    {
        // 1. Create a sample DOCX document.
        Document originalDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(originalDoc);
        builder.Writeln("Hello world with round‑trip info.");
        const string docxPath = "sample.docx";
        originalDoc.Save(docxPath, SaveFormat.Docx);

        if (!File.Exists(docxPath))
            throw new InvalidOperationException($"Failed to create '{docxPath}'.");

        // 2. Load the DOCX and save to HTML, enabling ExportRoundTripInfo.
        Document docForHtml = new Document(docxPath);
        HtmlSaveOptions htmlOptions = new HtmlSaveOptions(SaveFormat.Html);

        // The ExportRoundTripInfo property may not exist in older versions, so set it via reflection if present.
        PropertyInfo? exportProp = typeof(HtmlSaveOptions).GetProperty(
            "ExportRoundTripInfo",
            BindingFlags.Public | BindingFlags.Instance);

        if (exportProp != null && exportProp.CanWrite)
        {
            exportProp.SetValue(htmlOptions, true);
        }

        const string htmlPath = "output.html";
        docForHtml.Save(htmlPath, htmlOptions);

        if (!File.Exists(htmlPath))
            throw new InvalidOperationException($"Failed to create '{htmlPath}'.");

        // 3. Load the HTML back and save to a new DOCX to verify round‑trip.
        Document roundTripDoc = new Document(htmlPath);
        const string roundTripDocxPath = "roundtrip.docx";
        roundTripDoc.Save(roundTripDocxPath, SaveFormat.Docx);

        if (!File.Exists(roundTripDocxPath))
            throw new InvalidOperationException($"Failed to create '{roundTripDocxPath}'.");

        Console.WriteLine("Conversion completed successfully.");
    }
}
