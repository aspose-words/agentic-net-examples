using System;
using System.IO;
using System.Text.Json;
using Aspose.Words;
using Aspose.Words.Lists;

class InstructionManualGenerator
{
    static void Main()
    {
        // Sample JSON array of steps.
        // In a real scenario replace this string with the actual JSON source (e.g., read from a file).
        string stepsJson = @"[
            ""Open the application."",
            ""Log in with your credentials."",
            ""Navigate to the dashboard."",
            ""Select the report you wish to generate."",
            ""Click the 'Generate' button."",
            ""Save the report to your desired location.""
        ]";

        // Deserialize the JSON array into a string array.
        string[] steps = JsonSerializer.Deserialize<string[]>(stepsJson);

        // Create a new empty document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Add a numbered list based on the default numbered template (1., 2., 3., ...).
        List numberedList = doc.Lists.Add(ListTemplate.NumberDefault);

        // Apply the list to the builder.
        builder.ListFormat.List = numberedList;
        // Ensure we start at the first level (0‑based index).
        builder.ListFormat.ListLevelNumber = 0;

        // Write each step as a separate list item.
        foreach (string step in steps)
        {
            builder.Writeln(step);
        }

        // Remove list formatting from any subsequent paragraphs (optional cleanup).
        builder.ListFormat.RemoveNumbers();

        // Save the document to a file.
        string outputPath = Path.Combine(Environment.CurrentDirectory, "InstructionManual.docx");
        doc.Save(outputPath);
        Console.WriteLine($"Instruction manual saved to: {outputPath}");
    }
}
