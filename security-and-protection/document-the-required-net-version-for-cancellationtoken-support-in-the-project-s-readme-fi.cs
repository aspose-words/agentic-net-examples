using System;
using System.IO;

public class Program
{
    public static void Main()
    {
        // Required .NET version for CancellationToken support.
        const string requiredVersion = "net6.0";

        // Path to the README file.
        string readmePath = "README.md";

        // Content describing the .NET version requirement.
        string content = $"# Project README\n\n" +
                         $"## .NET Version Requirement\n\n" +
                         $"This project requires **{requiredVersion}** or later to support `CancellationToken` functionality.\n";

        // Write the README file (overwrites if it already exists).
        File.WriteAllText(readmePath, content);

        // Verify that the file was created successfully.
        if (!File.Exists(readmePath))
        {
            throw new InvalidOperationException("Failed to create README.md file.");
        }

        // Inform the user that the operation succeeded.
        Console.WriteLine($"README.md created with required .NET version: {requiredVersion}");
    }
}
