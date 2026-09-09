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

        // Step 1: Create a sample document that uses a font that likely does not exist.
        string sourceDocPath = Path.Combine(outputDir, "MissingFont.docx");
        Document tempDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(tempDoc);
        builder.Font.Name = "MissingFont";
        builder.Writeln("This text is formatted with a missing font.");
        tempDoc.Save(sourceDocPath);

        // Step 2: Configure FontSettings to substitute missing fonts with a known font (e.g., Arial).
        FontSettings fontSettings = new FontSettings();
        fontSettings.SubstitutionSettings.DefaultFontSubstitution.DefaultFontName = "Arial";

        // Step 3: Create LoadOptions and assign the FontSettings.
        LoadOptions loadOptions = new LoadOptions();
        loadOptions.FontSettings = fontSettings;

        // Step 4: Load the document using the configured LoadOptions.
        Document doc = new Document(sourceDocPath, loadOptions);

        // Optional: Keep original font metrics after substitution.
        doc.LayoutOptions.KeepOriginalFontMetrics = true;

        // Step 5: Attach a warning callback to capture any font substitution warnings.
        WarningInfoCollection warningCollector = new WarningInfoCollection();
        doc.WarningCallback = warningCollector;

        // Step 6: Save the loaded document to PDF; missing fonts will be substituted automatically.
        string pdfPath = Path.Combine(outputDir, "Result.pdf");
        doc.Save(pdfPath, SaveFormat.Pdf);

        // Output any captured warnings to the console.
        foreach (WarningInfo info in warningCollector)
        {
            if (info.WarningType == WarningType.FontSubstitution)
                Console.WriteLine(info.Description);
        }

        // Indicate completion.
        Console.WriteLine($"Document saved to: {pdfPath}");
    }
}
