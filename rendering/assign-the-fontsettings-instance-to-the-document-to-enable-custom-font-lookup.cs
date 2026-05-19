using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Fonts;

public class Program
{
    public static void Main()
    {
        // Define folders for output files.
        string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "Output");
        Directory.CreateDirectory(outputDir);

        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Write a paragraph using a font that is unlikely to be installed.
        builder.Font.Name = "NonExistentFont";
        builder.Writeln("This text uses a missing font and will be substituted.");

        // Create a FontSettings instance and assign it to the document.
        FontSettings fontSettings = new FontSettings();

        // Configure a simple substitution rule: use Arial when the requested font is missing.
        fontSettings.SubstitutionSettings.TableSubstitution.AddSubstitutes("NonExistentFont", new[] { "Arial" });

        // Assign the FontSettings to the document.
        doc.FontSettings = fontSettings;

        // Save the document to PDF to trigger layout and font resolution.
        string pdfPath = Path.Combine(outputDir, "Result.pdf");
        doc.Save(pdfPath, SaveFormat.Pdf);

        // Verify that the file was created.
        if (!File.Exists(pdfPath))
            throw new FileNotFoundException("The PDF file was not created.", pdfPath);

        // Optionally, write a short confirmation to the console.
        Console.WriteLine($"Document saved successfully to: {pdfPath}");
    }
}
