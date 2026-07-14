using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Loading;
using Aspose.Words.Fonts;

public class Program
{
    public static void Main()
    {
        // Define folders for input and output.
        string artifactsDir = Path.Combine(Directory.GetCurrentDirectory(), "Artifacts");
        Directory.CreateDirectory(artifactsDir);

        // -----------------------------------------------------------------
        // Step 1: Create a sample document that uses a font that is unlikely
        // to exist on the system ("MissingFont").
        // -----------------------------------------------------------------
        Document sampleDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(sampleDoc);
        builder.Font.Name = "MissingFont";
        builder.Writeln("This text is formatted with a missing font.");

        string sourceDocPath = Path.Combine(artifactsDir, "MissingFont.docx");
        sampleDoc.Save(sourceDocPath);

        // -----------------------------------------------------------------
        // Step 2: Configure FontSettings to substitute missing fonts.
        // We set the default substitution font to a common font (Arial).
        // -----------------------------------------------------------------
        FontSettings fontSettings = new FontSettings();
        fontSettings.SubstitutionSettings.DefaultFontSubstitution.DefaultFontName = "Arial";
        // Enable the default substitution rule (enabled by default, set explicitly for clarity).
        fontSettings.SubstitutionSettings.DefaultFontSubstitution.Enabled = true;

        // -----------------------------------------------------------------
        // Step 3: Load the document with the configured FontSettings.
        // -----------------------------------------------------------------
        LoadOptions loadOptions = new LoadOptions();
        loadOptions.FontSettings = fontSettings;

        Document loadedDoc = new Document(sourceDocPath, loadOptions);

        // -----------------------------------------------------------------
        // Step 4: Save the loaded document to PDF.
        // The missing font will be automatically substituted with Arial.
        // -----------------------------------------------------------------
        string outputPdfPath = Path.Combine(artifactsDir, "Result.pdf");
        loadedDoc.Save(outputPdfPath);
    }
}
