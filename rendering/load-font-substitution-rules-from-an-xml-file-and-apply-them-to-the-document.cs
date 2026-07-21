using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Fonts;

public class Program
{
    public static void Main()
    {
        // Prepare output folder.
        string artifactsDir = Path.Combine(Directory.GetCurrentDirectory(), "Artifacts");
        Directory.CreateDirectory(artifactsDir);

        // -----------------------------------------------------------------
        // 1. Create a sample font substitution table and save it to XML.
        // -----------------------------------------------------------------
        string xmlPath = Path.Combine(artifactsDir, "FontSubstitutionRules.xml");
        FontSettings tempSettings = new FontSettings();
        TableSubstitutionRule tempRule = tempSettings.SubstitutionSettings.TableSubstitution;
        // Define a substitution: when "MissingFont" is not found, use "Arial" then "Times New Roman".
        tempRule.AddSubstitutes("MissingFont", "Arial", "Times New Roman");
        // Save the table to an XML file.
        tempRule.Save(xmlPath);

        // -----------------------------------------------------------------
        // 2. Build a simple document that uses a font not present on the system.
        // -----------------------------------------------------------------
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Font.Name = "MissingFont";
        builder.Writeln("This line is formatted with a font that does not exist and should be substituted.");

        // -----------------------------------------------------------------
        // 3. Load the substitution rules from the XML file and apply them.
        // -----------------------------------------------------------------
        FontSettings fontSettings = new FontSettings();
        TableSubstitutionRule substitutionRule = fontSettings.SubstitutionSettings.TableSubstitution;
        substitutionRule.Load(xmlPath);
        doc.FontSettings = fontSettings;

        // -----------------------------------------------------------------
        // 4. Render the document to PDF.
        // -----------------------------------------------------------------
        string pdfPath = Path.Combine(artifactsDir, "Result.pdf");
        doc.Save(pdfPath);

        // -----------------------------------------------------------------
        // 5. Verify that the PDF was created.
        // -----------------------------------------------------------------
        if (!File.Exists(pdfPath))
            throw new InvalidOperationException("The PDF file was not created.");

        Console.WriteLine("PDF generated successfully at: " + pdfPath);
    }
}
