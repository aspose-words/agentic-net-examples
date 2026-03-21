using System;
using System.Collections.Generic;
using System.IO;
using System.Text;

class TrackRevisionsToggle
{
    static void Main(string[] args)
    {
        // Simulated document content.
        var lines = new List<string>();
        bool tracking = false;

        // Initial content.
        lines.Add("Initial content.");

        // If commands are supplied via arguments, process them.
        // Otherwise, show usage and exit (avoids waiting for interactive input).
        if (args.Length == 0)
        {
            Console.WriteLine("Usage: dotnet run -- <command1> <command2> ...");
            Console.WriteLine("Commands: on, off, exit");
            Console.WriteLine("Example: dotnet run -- on off exit");
            // Save the document as‑is and finish.
            SaveDocument(lines);
            return;
        }

        foreach (var raw in args)
        {
            var input = raw.Trim().ToLowerInvariant();

            if (input == "exit")
                break;

            if (input == "on")
            {
                tracking = true;
                lines.Add("[Tracked] Added while tracking is ON.");
                Console.WriteLine("Tracking started. Added a line.");
            }
            else if (input == "off")
            {
                tracking = false;
                lines.Add("[Tracked] Added while tracking is OFF.");
                Console.WriteLine("Tracking stopped. Added a line.");
            }
            else if (!string.IsNullOrWhiteSpace(input))
            {
                Console.WriteLine($"Unrecognized command '{input}'. Use 'on', 'off', or 'exit'.");
            }
        }

        SaveDocument(lines);
    }

    private static void SaveDocument(List<string> lines)
    {
        string outputPath = "TrackedDocument.txt";
        File.WriteAllLines(outputPath, lines, Encoding.UTF8);
        Console.WriteLine($"Document saved as '{outputPath}'.");
    }
}
