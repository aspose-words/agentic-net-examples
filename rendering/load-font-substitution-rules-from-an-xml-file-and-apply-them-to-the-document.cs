using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Fonts;

public class Program
{
    public static void Main()
    {
        // Paths for temporary files
        string xmlPath = "fontSubstitutions.xml";
        string docPath = "sample.docx";
        string pdfPath = "output.pdf";

        // 1. Create an XML file that defines font substitution rules.
        //    Map a non‑existent font ("NonExistentFont") to a common system font ("Arial").
        string xmlContent = @"<?xml version=""1.0"" encoding=""utf-8""?>
<FontSubstitutes>
    <Substitute>
        <Original>NonExistentFont</Original>
        <Substitute>Arial</Substitute>
    </Substitute>
</FontSubstitutes>";
        File.WriteAllText(xmlPath, xmlContent);

        // 2. Build a sample document that uses the missing font.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Font.Name = "NonExistentFont";
        builder.Writeln("This paragraph uses a font that does not exist on the system.");
        doc.Save(docPath);

        // 3. Load the font substitution rules from the XML file.
        FontSettings fontSettings = new FontSettings();
        fontSettings.SubstitutionSettings.TableSubstitution.Load(xmlPath);

        // 4. Apply the FontSettings to the document.
        doc.FontSettings = fontSettings;

        // 5. Render the document to PDF (the missing font should be substituted with Arial).
        doc.Save(pdfPath, SaveFormat.Pdf);

        // 6. Validate that the PDF file was created.
        if (!File.Exists(pdfPath))
            throw new Exception("PDF rendering failed – output file not found.");

        // 7. Simple verification: check the PDF content for a font substitution marker.
        //    Subset fonts are usually indicated by a six‑letter prefix followed by '+'.
        string pdfText = File.ReadAllText(pdfPath);
        bool containsSubsetMarker = pdfText.Contains("+");
        if (!containsSubsetMarker)
            throw new Exception("Font substitution may not have been applied – no subset font marker found.");

        // Cleanup (optional): delete temporary files if desired.
        // File.Delete(xmlPath);
        // File.Delete(docPath);
        // File.Delete(pdfPath);
    }
}
