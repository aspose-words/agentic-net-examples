using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Loading;
using Aspose.Words.Fonts;

public class Program
{
    public static void Main()
    {
        // Prepare output folder.
        string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "Output");
        Directory.CreateDirectory(outputDir);

        // Paths for the temporary source document and the final PDF.
        string sourceDocPath = Path.Combine(outputDir, "Source.docx");
        string resultPdfPath = Path.Combine(outputDir, "Result.pdf");

        // -----------------------------------------------------------------
        // 1. Create a document that uses a font which is unlikely to exist.
        // -----------------------------------------------------------------
        Document tempDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(tempDoc);
        builder.Font.Name = "MissingFontXYZ"; // Intentionally missing font.
        builder.Writeln("This text is formatted with a missing font.");

        // Save the temporary document.
        tempDoc.Save(sourceDocPath);

        // ---------------------------------------------------------------
        // 2. Configure FontSettings to substitute missing fonts automatically.
        // ---------------------------------------------------------------
        FontSettings fontSettings = new FontSettings();

        // Use a default substitute font that is expected to be present on the system.
        fontSettings.SubstitutionSettings.DefaultFontSubstitution.DefaultFontName = "Arial";

        // Enable the default font substitution rule (enabled by default, set explicitly for clarity).
        fontSettings.SubstitutionSettings.DefaultFontSubstitution.Enabled = true;

        // Optionally enable FontInfo substitution to improve matching.
        fontSettings.SubstitutionSettings.FontInfoSubstitution.Enabled = true;

        // ---------------------------------------------------------------
        // 3. Load the document with the configured FontSettings.
        // ---------------------------------------------------------------
        LoadOptions loadOptions = new LoadOptions
        {
            FontSettings = fontSettings
        };

        Document loadedDoc = new Document(sourceDocPath, loadOptions);

        // Preserve original font metrics after substitution (optional, default is true).
        loadedDoc.LayoutOptions.KeepOriginalFontMetrics = true;

        // ---------------------------------------------------------------
        // 4. Save the loaded document to PDF – missing fonts will be substituted.
        // ---------------------------------------------------------------
        loadedDoc.Save(resultPdfPath);

        // Verify that the output file was created.
        if (File.Exists(resultPdfPath))
        {
            Console.WriteLine("PDF generated successfully at: " + resultPdfPath);
        }
        else
        {
            Console.WriteLine("Failed to generate PDF.");
        }
    }
}
