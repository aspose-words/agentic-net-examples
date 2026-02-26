using System;
using System.IO;
using System.Linq;
using Aspose.Words;
using System.Drawing;

class Program
{
    static void Main()
    {
        // Load the DOTX template document.
        Document doc = new Document("Template.dotx");

        // Use LINQ to find all Run nodes whose font size is greater than 12 points.
        var largeRuns = doc.GetChildNodes(NodeType.Run, true)
                           .Cast<Run>()
                           .Where(r => r.Font.Size > 12)
                           .ToList();

        // Example action: output the number of matching runs.
        Console.WriteLine($"Runs with font size > 12 points: {largeRuns.Count}");

        // Optional: modify the found runs (e.g., change their color to red).
        foreach (var run in largeRuns)
        {
            run.Font.Color = Color.Red;
        }

        // Save the resulting document.
        doc.Save("Result.docx");
    }
}
