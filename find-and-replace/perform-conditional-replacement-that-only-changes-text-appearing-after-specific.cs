using System;
using System.IO;
using System.Text.RegularExpressions;

class ReplaceAfterHeading
{
    static void Main()
    {
        const string inputPath = "Input.txt";
        const string outputPath = "Output.txt";
        const string headingText = "My Heading";

        if (!File.Exists(inputPath))
        {
            Console.Error.WriteLine($"Input file \"{inputPath}\" not found.");
            return;
        }

        string[] lines = File.ReadAllLines(inputPath);
        int headingIndex = Array.FindIndex(lines, line => line.Trim() == headingText);

        if (headingIndex < 0)
        {
            Console.Error.WriteLine("Heading not found.");
            return;
        }

        var regex = new Regex(@"\bfoo\b", RegexOptions.IgnoreCase);
        for (int i = headingIndex + 1; i < lines.Length; i++)
        {
            lines[i] = regex.Replace(lines[i], "bar");
        }

        File.WriteAllLines(outputPath, lines);
        Console.WriteLine($"Processed file saved to \"{outputPath}\".");
    }
}
