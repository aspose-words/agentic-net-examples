using System;
using System.IO;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Fonts;

public class Program
{
    public static void Main()
    {
        // Create a new empty document.
        Document doc = new Document();

        // Initialize FontSettings and assign to the document.
        FontSettings fontSettings = new FontSettings();
        doc.FontSettings = fontSettings;

        // Define a substitution rule: replace missing "Times New Roman" with "Calibri".
        fontSettings.SubstitutionSettings.TableSubstitution.SetSubstitutes("Times New Roman", "Calibri");

        // Validate that the substitution rule was set correctly.
        var substitutes = fontSettings.SubstitutionSettings.TableSubstitution.GetSubstitutes("Times New Roman");
        if (substitutes == null || !substitutes.Contains("Calibri"))
            throw new InvalidOperationException("Failed to set the font substitution rule.");

        // Write some text using the missing font.
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Font.Name = "Times New Roman";
        builder.Writeln("This line uses Times New Roman, which will be substituted with Calibri.");

        // Save the document to a PDF file.
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "CustomFontSubstitution.pdf");
        doc.Save(outputPath);

        // Ensure the output file exists.
        if (!File.Exists(outputPath))
            throw new FileNotFoundException("The output file was not created.", outputPath);
    }
}
