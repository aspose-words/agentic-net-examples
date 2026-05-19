using System;
using System.IO;

public class Program
{
    public static void Main()
    {
        const string readmePath = "README.md";
        string content = "# Project README\n\n" +
                         "## .NET Version Requirement\n\n" +
                         "This project requires **.NET 6.0** or later to use `CancellationToken` features.\n";

        File.WriteAllText(readmePath, content);
        Console.WriteLine($"README file created at: {Path.GetFullPath(readmePath)}");
    }
}
