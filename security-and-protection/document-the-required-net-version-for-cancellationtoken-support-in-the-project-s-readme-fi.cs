using System;
using System.IO;

public class Program
{
    public static void Main()
    {
        // Path to the README file that will be created in the current directory.
        string readmePath = "README.md";

        // Content describing the .NET versions that support CancellationToken.
        string content = @"# Project README

## .NET Version Requirement for CancellationToken

The `CancellationToken` struct is available in the following .NET versions:

- .NET Framework 4.5 and later
- .NET Core 2.0 and later
- .NET 5.0 and later (including .NET 6, .NET 7, etc.)

Ensure your project targets one of these versions to use `CancellationToken` in asynchronous operations.
";

        // Write the content to the README file.
        File.WriteAllText(readmePath, content);

        // Validate that the file was created successfully.
        if (!File.Exists(readmePath))
            throw new InvalidOperationException("Failed to create README.md file.");

        // Inform the user (optional, does not require input).
        Console.WriteLine($"README.md has been created at: {Path.GetFullPath(readmePath)}");
    }
}
