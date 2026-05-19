using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Fonts;

public class Program
{
    public static void Main()
    {
        // Paths for temporary XML and output PDF.
        string xmlPath = Path.Combine(Directory.GetCurrentDirectory(), "FontSubstitutionRules.xml");
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "CustomFontSubstitution.pdf");

        // -----------------------------------------------------------------
        // Step 1: Create a font substitution table using the built‑in Windows settings
        // and save it to an XML file.
        // -----------------------------------------------------------------
        FontSettings tempSettings = new FontSettings();
        TableSubstitutionRule tempTable = tempSettings.SubstitutionSettings.TableSubstitution;
        tempTable.LoadWindowsSettings();               // Load default Windows substitution table.
        tempTable.Save(xmlPath);                       // Persist the table to XML.

        // -----------------------------------------------------------------
        // Step 2: Load the custom substitution table from the XML file.
        // -----------------------------------------------------------------
        FontSettings fontSettings = new FontSettings();
        TableSubstitutionRule table = fontSettings.SubstitutionSettings.TableSubstitution;
        table.Load(xmlPath);                           // Load the previously saved table.

        // -----------------------------------------------------------------
        // Step 3: Create a document that uses a font not present on the system.
        // The substitution table will provide a fallback font.
        // -----------------------------------------------------------------
        Document doc = new Document();
        doc.FontSettings = fontSettings;               // Apply the custom font settings to the document.

        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Font.Name = "Amethysta";               // Font likely unavailable on the machine.
        builder.Writeln("This line is formatted with the missing font \"Amethysta\".");
        builder.Font.Name = "Arial";
        builder.Writeln("This line uses a standard font and will render normally.");

        // -----------------------------------------------------------------
        // Step 4: Save the resulting document.
        // -----------------------------------------------------------------
        doc.Save(outputPath);
    }
}
