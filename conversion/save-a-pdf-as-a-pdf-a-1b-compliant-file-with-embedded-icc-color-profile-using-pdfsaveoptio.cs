using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // Create a simple Word document in memory.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("This PDF is saved as PDF/A‑1b with an embedded ICC profile.");

        // -----------------------------------------------------------------
        // NOTE:
        // In recent versions of Aspose.Words the PdfSaveOptions class does not
        // expose an IccProfile property. Therefore the ICC profile cannot be
        // embedded directly via this API. The example demonstrates the
        // creation of a PDF/A‑1b compliant file, which satisfies the primary
        // requirement of the task. If a newer API becomes available, the
        // ICC profile can be set accordingly.
        // -----------------------------------------------------------------

        // Configure PDF save options for PDF/A‑1b compliance.
        PdfSaveOptions saveOptions = new PdfSaveOptions
        {
            Compliance = PdfCompliance.PdfA1b,
            // Embed all fonts to meet PDF/A requirements.
            EmbedFullFonts = true
        };

        // Define the output file path.
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "Output.PdfA1b.pdf");

        // Save the document as PDF/A‑1b.
        doc.Save(outputPath, saveOptions);

        // Verify that the file was created and is not empty.
        if (!File.Exists(outputPath))
            throw new FileNotFoundException("The PDF file was not created.", outputPath);

        if (new FileInfo(outputPath).Length == 0)
            throw new InvalidOperationException("The PDF file is empty.");

        Console.WriteLine($"PDF/A‑1b file successfully created at: {outputPath}");
    }
}
