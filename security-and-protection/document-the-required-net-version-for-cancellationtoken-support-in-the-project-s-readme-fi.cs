using System;
using System.IO;

public class Program
{
    public static void Main()
    {
        // Define the required .NET version for CancellationToken support.
        // CancellationToken was introduced in .NET Framework 4.5,
        // .NET Core 2.0 and is also available in .NET Standard 2.0.
        string readmeContent = @"# Project README

## CancellationToken Support

CancellationToken is supported starting from:
- **.NET Framework 4.5**
- **.NET Core 2.0**
- **.NET Standard 2.0**

Ensure your project targets one of the above frameworks or a later version.";

        // Write the content to a README.md file in the current directory.
        const string fileName = "README.md";
        File.WriteAllText(fileName, readmeContent);

        // Inform the user that the file has been created.
        Console.WriteLine($"README file '{fileName}' has been created with the required .NET version information.");
    }
}
