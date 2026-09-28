using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // Create a simple Word document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("This PDF is saved as PDF/A‑1b.");

        // Configure PDF save options for PDF/A‑1b compliance.
        PdfSaveOptions saveOptions = new PdfSaveOptions
        {
            Compliance = PdfCompliance.PdfA1b
            // The current Aspose.Words version does not expose EmbedColorProfile or IccProfile properties.
            // PDF/A‑1b compliance is ensured; embedding an ICC profile would require a newer API.
        };

        // Save the document as PDF/A‑1b.
        const string outputPath = "output.pdf";
        doc.Save(outputPath, saveOptions);

        // Verify that the file was created.
        if (!File.Exists(outputPath))
            throw new InvalidOperationException("The PDF/A‑1b file was not created.");

        // Verify that the file size is greater than zero.
        FileInfo info = new FileInfo(outputPath);
        if (info.Length == 0)
            throw new InvalidOperationException("The PDF/A‑1b file is empty.");
    }

    // Returns a byte array containing an ICC profile.
    // Placeholder method retained for completeness; not used in this example.
    private static byte[] GetIccProfile()
    {
        // A very small dummy ICC profile (not a real color profile).
        // Replace with a valid ICC profile byte array for production use.
        return new byte[] { 0x00, 0x01, 0x02, 0x03 };
    }
}
