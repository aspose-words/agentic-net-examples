using System;
using System.IO;

public class Program
{
    public static void Main()
    {
        // Path to the README file that will be created/overwritten.
        const string readmePath = "README.md";

        // Content describing the .NET version required for CancellationToken support.
        string content = "# Project README\n\n" +
                         "## .NET Version Requirement\n\n" +
                         "This project requires **.NET 5.0** or later to support `CancellationToken`.\n";

        // Write the content to the README file.
        File.WriteAllText(readmePath, content);

        // Inform the user where the file was created.
        Console.WriteLine($"README file created at: {Path.GetFullPath(readmePath)}");
    }
}
