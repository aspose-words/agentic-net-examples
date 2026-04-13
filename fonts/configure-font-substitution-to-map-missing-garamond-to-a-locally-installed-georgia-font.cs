using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Fonts;

public class Program
{
    public static void Main()
    {
        // Create a new empty document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Write a line using a font that is likely not installed (Garamond).
        builder.Font.Name = "Garamond";
        builder.Writeln("This text uses Garamond, which will be substituted with Georgia.");

        // Validate that the font name was set correctly.
        if (builder.Font.Name != "Garamond")
            throw new InvalidOperationException("Failed to set the initial font name.");

        // Configure font substitution: map missing Garamond to the locally installed Georgia font.
        FontSettings fontSettings = new FontSettings();
        // Use the table substitution rule to define the substitute.
        fontSettings.SubstitutionSettings.TableSubstitution.SetSubstitutes("Garamond", "Georgia");

        // Validate that the substitution rule was recorded.
        bool hasGeorgia = false;
        foreach (string substitute in fontSettings.SubstitutionSettings.TableSubstitution.GetSubstitutes("Garamond"))
        {
            if (substitute == "Georgia")
            {
                hasGeorgia = true;
                break;
            }
        }
        if (!hasGeorgia)
            throw new InvalidOperationException("The substitution for Garamond was not set correctly.");

        // Apply the font settings to the document.
        doc.FontSettings = fontSettings;

        // Save the document to PDF (the missing Garamond will be rendered using Georgia).
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "FontSubstitutionExample.pdf");
        doc.Save(outputPath);

        // Ensure the output file was created.
        if (!File.Exists(outputPath))
            throw new FileNotFoundException("The output PDF was not created.", outputPath);
    }
}
