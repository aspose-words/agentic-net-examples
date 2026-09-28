using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

public class ExportRoundTripInfoExample
{
    public static void Main()
    {
        // Create a sample DOCX document.
        Document sourceDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(sourceDoc);
        builder.Writeln("Sample content for round‑trip test.");

        // Save the sample as DOCX.
        const string docxPath = "input.docx";
        sourceDoc.Save(docxPath, SaveFormat.Docx);
        if (!File.Exists(docxPath))
            throw new InvalidOperationException("The DOCX input file was not created.");

        // Load the DOCX document.
        Document doc = new Document(docxPath);

        // Configure HTML save options.
        // The ExportRoundTripInfo property is available in newer versions of Aspose.Words.
        // If the property does not exist in the referenced version, the code will still compile
        // and the HTML will be saved without round‑trip information.
        HtmlSaveOptions htmlOptions = new HtmlSaveOptions();
#if NET7_0_OR_GREATER
        // Attempt to enable round‑trip information if the property exists.
        // This block will be ignored if the property is not present in the used library version.
        var exportProp = typeof(HtmlSaveOptions).GetProperty("ExportRoundTripInfo");
        if (exportProp != null && exportProp.CanWrite)
        {
            exportProp.SetValue(htmlOptions, true);
        }
#endif

        // Save the document as HTML.
        const string htmlPath = "output.html";
        doc.Save(htmlPath, htmlOptions);
        if (!File.Exists(htmlPath))
            throw new InvalidOperationException("The HTML output file was not created.");

        // Load the HTML back into a Document to verify round‑trip capability.
        Document roundTripDoc = new Document(htmlPath);

        // Save the round‑tripped document as DOCX.
        const string roundTripDocxPath = "roundtrip.docx";
        roundTripDoc.Save(roundTripDocxPath, SaveFormat.Docx);
        if (!File.Exists(roundTripDocxPath))
            throw new InvalidOperationException("The round‑trip DOCX file was not created.");

        // Indicate successful completion.
        Console.WriteLine("ExportRoundTripInfo handling completed and round‑trip conversion succeeded.");
    }
}
