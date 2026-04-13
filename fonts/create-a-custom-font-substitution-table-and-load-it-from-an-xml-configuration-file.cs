using System;
using System.IO;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Fonts;

public class Program
{
    public static void Main()
    {
        // Path for the custom font substitution XML file.
        const string xmlPath = "font_substitution.xml";

        // Simple XML defining a substitute for a missing font.
        const string xmlContent = @"<?xml version=""1.0"" encoding=""utf-8""?>
<substitution>
    <font name=""MissingFont"">
        <substitute>Times New Roman</substitute>
    </font>
</substitution>";

        // Write the XML configuration to disk.
        File.WriteAllText(xmlPath, xmlContent);

        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Use a font that does not exist on the system.
        builder.Font.Name = "MissingFont";
        builder.Writeln("This text uses a missing font and should be substituted according to the custom table.");

        // Set up FontSettings for the document.
        FontSettings fontSettings = new FontSettings();
        doc.FontSettings = fontSettings;

        // Load the custom substitution table from the XML file.
        TableSubstitutionRule tableRule = fontSettings.SubstitutionSettings.TableSubstitution;
        tableRule.Load(xmlPath);

        // Validate that the substitution was loaded.
        var substitutes = tableRule.GetSubstitutes("MissingFont");
        if (substitutes != null && substitutes.Any())
        {
            Console.WriteLine("Substitutes for 'MissingFont': " + string.Join(", ", substitutes));
        }
        else
        {
            Console.WriteLine("No substitutes found for 'MissingFont'.");
        }

        // Save the document to PDF.
        const string outputPath = "output.pdf";
        doc.Save(outputPath);

        // Verify that the output file was created.
        Console.WriteLine(File.Exists(outputPath)
            ? $"Document successfully saved to '{outputPath}'."
            : $"Failed to save document to '{outputPath}'.");
    }
}
