using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // Define output file names.
        string defaultPdf = "Default_OpenType.pdf";
        string disabledPdf = "Disabled_OpenType.pdf";

        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Add sample text that contains characters affected by OpenType features (e.g., ligatures).
        builder.Font.Name = "Times New Roman";
        builder.Writeln("OpenType ligature test: office, official, affix, flake, ﬁ, ﬂ.");

        // Save the document with default OpenType handling.
        doc.Save(defaultPdf, SaveFormat.Pdf);

        // Disable OpenType font formatting features via CompatibilityOptions.
        doc.CompatibilityOptions.DisableOpenTypeFontFormattingFeatures = true;

        // Save the document after disabling OpenType features.
        doc.Save(disabledPdf, SaveFormat.Pdf);

        // Verify that both PDF files were created successfully.
        if (!File.Exists(defaultPdf))
            throw new FileNotFoundException($"Failed to create {defaultPdf}");
        if (!File.Exists(disabledPdf))
            throw new FileNotFoundException($"Failed to create {disabledPdf}");

        Console.WriteLine($"PDF files generated: {defaultPdf}, {disabledPdf}");
    }
}
