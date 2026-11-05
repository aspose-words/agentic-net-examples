using System;
using System.IO;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Define the font names to replace.
        string oldFontName = "Arial";
        string newFontName = "Times New Roman";

        // Create a new document and add sample text.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Paragraph with the old font.
        builder.Font.Name = oldFontName;
        builder.Writeln($"This line uses the old font: {oldFontName}.");

        // Paragraph with a different font that should remain unchanged.
        builder.Font.Name = "Calibri";
        builder.Writeln("This line uses Calibri and should stay unchanged.");

        // Another paragraph with the old font.
        builder.Font.Name = oldFontName;
        builder.Writeln($"Another occurrence of the old font: {oldFontName}.");

        // Replace all occurrences of the old font with the new font.
        foreach (Node node in doc.GetChildNodes(NodeType.Run, true))
        {
            Run run = (Run)node;
            if (!string.IsNullOrEmpty(run.Font.Name) &&
                run.Font.Name.Equals(oldFontName, StringComparison.OrdinalIgnoreCase))
            {
                run.Font.Name = newFontName;

                // Validate that the font was set correctly.
                if (!run.Font.Name.Equals(newFontName, StringComparison.OrdinalIgnoreCase))
                {
                    Console.WriteLine("Warning: Failed to set the new font on a run.");
                }
            }
        }

        // Save the modified document.
        string outputPath = "Output.docx";
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
