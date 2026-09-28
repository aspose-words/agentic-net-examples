using System;
using System.IO;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Create a new document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert a paragraph with a run of text.
        builder.Writeln("Sample text before formatting.");

        // Retrieve the first run in the document.
        Run run = (Run)doc.FirstSection.Body.Paragraphs[0].Runs[0];

        // Apply various font attributes.
        run.Font.Name = "Arial";
        run.Font.Size = 16;
        run.Font.Bold = true;
        run.Font.Color = System.Drawing.Color.Red;

        // Display font attributes before clearing.
        Console.WriteLine("Before ClearFormatting:");
        PrintFontInfo(run.Font);

        // Reset all font attributes to defaults.
        run.Font.ClearFormatting();

        // Display font attributes after clearing.
        Console.WriteLine("After ClearFormatting:");
        PrintFontInfo(run.Font);

        // Save the document.
        string outputPath = "ResetFontExample.docx";
        doc.Save(outputPath);

        // Verify that the file was created.
        Console.WriteLine(File.Exists(outputPath) ? "File saved successfully." : "Failed to save file.");
    }

    private static void PrintFontInfo(Aspose.Words.Font font)
    {
        // Helper method to output font properties.
        string name = string.IsNullOrEmpty(font.Name) ? "(default)" : font.Name;
        string size = font.Size > 0 ? font.Size.ToString() : "(default)";
        string bold = font.Bold ? "True" : "False";
        string color = font.Color.IsEmpty ? "(default)" : font.Color.ToString();

        Console.WriteLine($"Name: {name}, Size: {size}, Bold: {bold}, Color: {color}");
    }
}
