using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Fonts;

public class Program
{
    public static void Main()
    {
        // Prepare output directory.
        string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "Output");
        Directory.CreateDirectory(outputDir);

        // Paths for the XML rules file and the resulting PDF.
        string xmlPath = Path.Combine(outputDir, "FontSubstitutionRules.xml");
        string pdfPath = Path.Combine(outputDir, "Result.pdf");

        // Create a sample document that uses a font which is unlikely to exist.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Font.Name = "MissingFont";
        builder.Writeln("This text uses a missing font and should be substituted.");

        // Define a substitution rule and save it to an XML file.
        FontSettings fontSettings = new FontSettings();
        TableSubstitutionRule tableRule = fontSettings.SubstitutionSettings.TableSubstitution;
        tableRule.AddSubstitutes("MissingFont", "Arial");
        tableRule.Save(xmlPath);

        // Load the substitution rules from the XML file into a new FontSettings instance.
        FontSettings loadedSettings = new FontSettings();
        TableSubstitutionRule loadedTableRule = loadedSettings.SubstitutionSettings.TableSubstitution;
        loadedTableRule.Load(xmlPath);

        // Apply the loaded font settings to the document.
        doc.FontSettings = loadedSettings;

        // Render the document to PDF.
        doc.Save(pdfPath, SaveFormat.Pdf);

        // Verify that the PDF was created.
        if (!File.Exists(pdfPath))
            throw new Exception("PDF was not created.");

        // Indicate successful completion.
        Console.WriteLine("PDF saved to: " + pdfPath);
    }
}
