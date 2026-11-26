using System;
using System.IO;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Create a new document.
        Document doc = new Document();

        // Add a paragraph to the first (and only) section.
        Paragraph para = new Paragraph(doc);
        doc.FirstSection.Body.AppendChild(para);

        // Create a Run with some text.
        Run run = new Run(doc, "Sample text with custom font attributes.");
        para.AppendChild(run);

        // Apply custom font formatting.
        run.Font.Name = "Arial";
        run.Font.Size = 16;
        run.Font.Bold = true;
        run.Font.Italic = true;
        run.Font.Underline = Underline.Double;
        run.Font.Color = System.Drawing.Color.Blue; // System.Drawing.Color used explicitly.

        // Output font properties before clearing.
        Console.WriteLine("Before ClearFormatting:");
        Console.WriteLine($"Name: {run.Font.Name}");
        Console.WriteLine($"Size: {run.Font.Size}");
        Console.WriteLine($"Bold: {run.Font.Bold}");
        Console.WriteLine($"Italic: {run.Font.Italic}");
        Console.WriteLine($"Underline: {run.Font.Underline}");
        Console.WriteLine($"Color: {run.Font.Color}");

        // Reset all font attributes to defaults.
        run.Font.ClearFormatting();

        // Output font properties after clearing.
        Console.WriteLine("\nAfter ClearFormatting:");
        Console.WriteLine($"Name: {(run.Font.Name ?? "null")}");
        Console.WriteLine($"Size: {run.Font.Size}");
        Console.WriteLine($"Bold: {run.Font.Bold}");
        Console.WriteLine($"Italic: {run.Font.Italic}");
        Console.WriteLine($"Underline: {run.Font.Underline}");
        Console.WriteLine($"Color: {run.Font.Color}");

        // Save the document.
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "Output.docx");
        doc.Save(outputPath);

        // Verify that the file was created.
        if (File.Exists(outputPath))
        {
            Console.WriteLine($"\nDocument saved successfully to: {outputPath}");
        }
        else
        {
            Console.WriteLine("\nFailed to save the document.");
        }
    }
}
