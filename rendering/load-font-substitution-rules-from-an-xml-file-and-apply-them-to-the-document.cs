using System;
using System.IO;
using System.Linq;
using System.Xml.Linq;
using Aspose.Words;
using Aspose.Words.Fonts;

public class Program
{
    public static void Main()
    {
        // Paths for the sample document, XML rules, and output PDF.
        string docPath = "Sample.docx";
        string xmlPath = "FontSubstitution.xml";
        string pdfPath = "Result.pdf";

        // -----------------------------------------------------------------
        // 1. Create a sample DOCX that uses a font likely not installed.
        // -----------------------------------------------------------------
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Font.Name = "NonExistentFont";
        builder.Writeln("This paragraph uses a font that may not be installed on the system.");
        doc.Save(docPath);

        // -----------------------------------------------------------------
        // 2. Create an XML file that defines font substitution rules.
        // -----------------------------------------------------------------
        XDocument substitutionXml = new XDocument(
            new XElement("FontSubstitutions",
                new XElement("Substitution",
                    new XElement("Original", "NonExistentFont"),
                    new XElement("Substitutes",
                        new XElement("Font", "Arial")
                    )
                )
            )
        );
        substitutionXml.Save(xmlPath);

        // -----------------------------------------------------------------
        // 3. Load the XML substitution rules.
        // -----------------------------------------------------------------
        XDocument loadedXml = XDocument.Load(xmlPath);
        var substitutions = loadedXml.Root?
            .Elements("Substitution")
            .Select(s => new
            {
                Original = (string)s.Element("Original"),
                Substitutes = s.Element("Substitutes")?
                    .Elements("Font")
                    .Select(f => (string)f)
                    .ToArray() ?? Array.Empty<string>()
            })
            .Where(s => !string.IsNullOrEmpty(s.Original) && s.Substitutes.Length > 0)
            .ToList();

        // -----------------------------------------------------------------
        // 4. Apply the substitution rules to FontSettings.
        // -----------------------------------------------------------------
        FontSettings fontSettings = new FontSettings();

        // Get the table substitution object from the substitution settings.
        var table = fontSettings.SubstitutionSettings.TableSubstitution;

        foreach (var sub in substitutions)
        {
            // Add substitutes for each original font.
            table.AddSubstitutes(sub.Original, sub.Substitutes);
        }

        // Assign the configured FontSettings to the document.
        doc.FontSettings = fontSettings;

        // -----------------------------------------------------------------
        // 5. Render the document to PDF using the applied substitutions.
        // -----------------------------------------------------------------
        doc.Save(pdfPath, SaveFormat.Pdf);

        // -----------------------------------------------------------------
        // 6. Validate that the PDF was created.
        // -----------------------------------------------------------------
        if (!File.Exists(pdfPath))
        {
            throw new Exception("The PDF output file was not created.");
        }

        Console.WriteLine("PDF generated successfully with font substitution applied.");
    }
}
