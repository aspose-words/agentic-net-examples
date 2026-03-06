using System;
using System.IO;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Drawing;

class Program
{
    static void Main()
    {
        // Path to the source DOTX template.
        string inputPath = @"C:\Docs\Template.dotx";

        // Load the DOTX document.
        Document doc = new Document(inputPath);

        // Retrieve all Run nodes in the document.
        var runs = doc.GetChildNodes(NodeType.Run, true)
                      .Cast<Run>()
                      .Where(r => r.Font.Size > 12) // Lambda expression to filter runs with font size > 12 points.
                      .ToList();

        // Output information about the matching runs.
        Console.WriteLine($"Found {runs.Count} run(s) with font size greater than 12 points:");
        foreach (Run run in runs)
        {
            // Display the text of the run and its font size.
            Console.WriteLine($"- Text: \"{run.Text}\" | Size: {run.Font.Size} pt");
        }

        // Optionally, save the document (unchanged) to a new file.
        string outputPath = @"C:\Docs\ProcessedTemplate.dotx";
        doc.Save(outputPath);
    }
}
