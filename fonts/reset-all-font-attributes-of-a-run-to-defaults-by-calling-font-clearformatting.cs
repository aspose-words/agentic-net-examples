using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Fonts;
using Aspose.Drawing; // For Aspose.Drawing.Color

public class Program
{
    public static void Main()
    {
        // Create a new empty document.
        Document doc = new Document();

        // Use DocumentBuilder to add a paragraph with a Run.
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("This is a sample run with custom font settings.");

        // Retrieve the Run that was just added.
        // After Writeln, the builder's current paragraph is a new empty paragraph,
        // so we need to get the previous paragraph which contains the run.
        Paragraph previousParagraph = builder.CurrentParagraph.PreviousSibling as Paragraph;
        if (previousParagraph == null || previousParagraph.Runs.Count == 0)
        {
            Console.WriteLine("Failed to locate the run.");
            return;
        }

        Run run = previousParagraph.Runs[0];

        // Apply some custom font formatting.
        run.Font.Name = "Arial";
        run.Font.Size = 16;
        run.Font.Bold = true;
        run.Font.Italic = true;
        run.Font.Underline = Underline.Single;

        // Use Aspose.Drawing.Color for creation and convert to System.Drawing.Color.
        Aspose.Drawing.Color asposeColor = Aspose.Drawing.Color.Black;
        run.Font.Color = System.Drawing.Color.FromArgb(asposeColor.ToArgb());

        // Reset all font attributes to defaults.
        run.Font.ClearFormatting();

        // Validate that the font properties have been reset to defaults.
        // Default values: Name = null, Size = 0, Bold = false, Italic = false, Underline = None, Color = Empty.
        Console.WriteLine("After ClearFormatting:");
        Console.WriteLine($"Name: {(run.Font.Name ?? "null")}");
        Console.WriteLine($"Size: {run.Font.Size}");
        Console.WriteLine($"Bold: {run.Font.Bold}");
        Console.WriteLine($"Italic: {run.Font.Italic}");
        Console.WriteLine($"Underline: {run.Font.Underline}");
        Console.WriteLine($"Color is empty: {run.Font.Color.IsEmpty}");

        // Save the document to disk.
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "Output.docx");
        doc.Save(outputPath);

        // Ensure the file was created.
        if (File.Exists(outputPath))
        {
            Console.WriteLine($"Document saved successfully to: {outputPath}");
        }
        else
        {
            Console.WriteLine("Failed to save the document.");
        }
    }
}
