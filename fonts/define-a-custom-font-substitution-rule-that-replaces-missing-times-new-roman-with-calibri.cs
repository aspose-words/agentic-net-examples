using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Fonts;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Set the font to Times New Roman (assumed to be missing on the system).
        builder.Font.Name = "Times New Roman";
        builder.Writeln("This text is originally set to Times New Roman.");

        // Configure font substitution: replace missing Times New Roman with Calibri.
        FontSettings fontSettings = new FontSettings();

        // The correct API for adding a substitution uses the TableSubstitution property.
        // It maps an unavailable font name to a list of fallback fonts.
        fontSettings.SubstitutionSettings.TableSubstitution.AddSubstitutes(
            "Times New Roman",
            new string[] { "Calibri" });

        // Apply the font settings to the document.
        doc.FontSettings = fontSettings;

        // Save the document to disk.
        string outputPath = "Output.docx";
        doc.Save(outputPath);

        // Verify that the file was created.
        if (File.Exists(outputPath))
        {
            Console.WriteLine("Document saved successfully.");
        }
        else
        {
            Console.WriteLine("Failed to save document.");
        }
    }
}
