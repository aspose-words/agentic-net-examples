using System;
using System.IO;
using System.Text;

class Program
{
    static void Main()
    {
        // Input and output file paths.
        const string inputPath = "Input.txt";
        const string outputPath = "Output.txt";

        // Ensure the input file exists; create a sample if it does not.
        if (!File.Exists(inputPath))
        {
            File.WriteAllLines(inputPath, new[]
            {
                "First paragraph.",
                "",
                "Second paragraph.",
                "Third paragraph."
            });
        }

        // Predefined equation representation (as plain text).
        const string equation = @"\f(1,2)";

        // Read all lines from the input file.
        string[] lines = File.ReadAllLines(inputPath, Encoding.UTF8);

        var sb = new StringBuilder();

        foreach (string line in lines)
        {
            // Write the original line.
            sb.AppendLine(line);

            // If the line is not empty or whitespace, insert the equation after it.
            if (!string.IsNullOrWhiteSpace(line))
            {
                sb.AppendLine(equation);
            }
        }

        // Write the processed content to the output file.
        File.WriteAllText(outputPath, sb.ToString(), Encoding.UTF8);

        Console.WriteLine($"Processing complete. Output written to '{outputPath}'.");
    }
}
