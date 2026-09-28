using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Fonts;

public class Program
{
    public static void Main()
    {
        // Create a new document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Set the font to Times New Roman (the font we will substitute).
        builder.Font.Name = "Times New Roman";
        builder.Writeln("This text should be rendered with Calibri because Times New Roman is missing.");

        // Configure a custom font substitution rule: replace Times New Roman with Calibri.
        FontSettings fontSettings = new FontSettings();
        fontSettings.SubstitutionSettings.TableSubstitution.AddSubstitutes("Times New Roman", new string[] { "Calibri" });
        doc.FontSettings = fontSettings;

        // Save the document.
        string outputPath = "CustomFontSubstitution.docx";
        doc.Save(outputPath);

        // Verify that the file was created.
        if (File.Exists(outputPath))
        {
            Console.WriteLine("Document saved successfully: " + Path.GetFullPath(outputPath));
        }
        else
        {
            Console.WriteLine("Failed to save the document.");
        }
    }
}
