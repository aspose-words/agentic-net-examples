using System;
using System.IO;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Fonts;
using Newtonsoft.Json;

public class Program
{
    public static void Main()
    {
        // Create a new document and add a paragraph with Garamond font.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Font.Name = "Garamond";
        builder.Writeln("This text should be rendered with Georgia if Garamond is not installed.");

        // Configure font substitution: map missing Garamond to Georgia.
        FontSettings fontSettings = new FontSettings();
        fontSettings.SubstitutionSettings.TableSubstitution.AddSubstitutes("Garamond", new[] { "Georgia" });
        doc.FontSettings = fontSettings;

        // Validate that the substitution was added.
        string[] substitutes = fontSettings.SubstitutionSettings.TableSubstitution
            .GetSubstitutes("Garamond")
            .ToArray();

        bool substitutionConfigured = substitutes != null &&
                                       substitutes.Length > 0 &&
                                       substitutes[0] == "Georgia";

        // Save the document.
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "Output.docx");
        doc.Save(outputPath);

        // Verify the file was created.
        bool fileExists = File.Exists(outputPath);

        // Output the result as JSON.
        var result = new
        {
            SubstitutionConfigured = substitutionConfigured,
            OutputFile = outputPath,
            FileExists = fileExists
        };
        string json = JsonConvert.SerializeObject(result, Formatting.Indented);
        Console.WriteLine(json);
    }
}
