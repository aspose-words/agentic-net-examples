using System;
using System.IO;

public class Program
{
    public static void Main()
    {
        // Define the path for the README file in the current directory.
        string readmePath = Path.Combine(Directory.GetCurrentDirectory(), "README.md");

        // Content describing the .NET version required for CancellationToken support.
        string content = "# Project README\n\n" +
                         "## .NET Version Requirement\n\n" +
                         "CancellationToken is supported starting from .NET Framework 4.0 and .NET Core 2.0 (and later versions). " +
                         "Ensure your project targets at least these versions.\n";

        // Write the content to the README file.
        File.WriteAllText(readmePath, content);

        // Validate that the file was created successfully.
        if (!File.Exists(readmePath))
        {
            throw new InvalidOperationException("Failed to create README.md file.");
        }
    }
}
