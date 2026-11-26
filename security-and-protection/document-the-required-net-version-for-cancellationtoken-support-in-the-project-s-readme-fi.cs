using System;
using System.IO;

public class Program
{
    public static void Main()
    {
        // Path for the README file
        string readmePath = "README.md";

        // Content describing the required .NET version for CancellationToken support
        string content = @"# CancellationToken Support

CancellationToken is available in the following .NET versions:

- **.NET Framework** 4.0 and later
- **.NET Standard** 1.0 and later
- **.NET Core** 1.0 and later
- **.NET 5** and later (including .NET 6, .NET 7, etc.)

These versions provide the `System.Threading.CancellationToken` struct and related APIs for cooperative cancellation in asynchronous and parallel operations.";

        // Write the content to the README file
        File.WriteAllText(readmePath, content);

        // Validate that the file was created
        if (!File.Exists(readmePath))
        {
            throw new InvalidOperationException($"Failed to create the README file at '{readmePath}'.");
        }

        // Optionally, output a confirmation (not required, but harmless)
        Console.WriteLine($"README file created at '{Path.GetFullPath(readmePath)}'.");
    }
}
