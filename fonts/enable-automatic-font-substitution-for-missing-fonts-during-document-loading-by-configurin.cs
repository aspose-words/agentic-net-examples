using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Fonts;
using Aspose.Words.Loading;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // Define output directories.
        string artifactsDir = Path.Combine(Directory.GetCurrentDirectory(), "Artifacts");
        Directory.CreateDirectory(artifactsDir);

        // -----------------------------------------------------------------
        // Step 1: Create a sample document that uses a font that does not exist.
        // -----------------------------------------------------------------
        Document docToSave = new Document();
        DocumentBuilder builder = new DocumentBuilder(docToSave);
        builder.Font.Name = "MissingFont"; // Intentionally missing.
        builder.Writeln("This text is formatted with a missing font.");

        string sourceDocPath = Path.Combine(artifactsDir, "MissingFont.docx");
        docToSave.Save(sourceDocPath);

        // -----------------------------------------------------------------
        // Step 2: Configure FontSettings to substitute missing fonts automatically.
        // -----------------------------------------------------------------
        FontSettings fontSettings = new FontSettings();
        // Use Arial as the fallback for any unavailable font.
        fontSettings.SubstitutionSettings.DefaultFontSubstitution.DefaultFontName = "Arial";
        // The DefaultFontSubstitution rule is enabled by default, but we set it explicitly for clarity.
        fontSettings.SubstitutionSettings.DefaultFontSubstitution.Enabled = true;

        // -----------------------------------------------------------------
        // Step 3: Load the document with the configured FontSettings.
        // -----------------------------------------------------------------
        LoadOptions loadOptions = new LoadOptions();
        loadOptions.FontSettings = fontSettings;

        Document loadedDoc = new Document(sourceDocPath, loadOptions);
        // Preserve original font metrics after substitution (optional).
        loadedDoc.LayoutOptions.KeepOriginalFontMetrics = true;

        // -----------------------------------------------------------------
        // Step 4: Save the loaded document to PDF. The missing font will be substituted.
        // -----------------------------------------------------------------
        string resultPdfPath = Path.Combine(artifactsDir, "Result.pdf");
        loadedDoc.Save(resultPdfPath, SaveFormat.Pdf);

        // Verify that the output file was created.
        if (File.Exists(resultPdfPath))
        {
            Console.WriteLine($"PDF saved successfully to: {resultPdfPath}");
        }
        else
        {
            Console.WriteLine("Failed to save PDF.");
        }
    }
}
