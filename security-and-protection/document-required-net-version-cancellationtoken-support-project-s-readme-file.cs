using System;
using System.IO;
using System.Threading;

namespace AsposeWordsDocumentation
{
    class Program
    {
        static void Main()
        {
            // The CancellationToken type was introduced in .NET Framework 4.0
            // and is also available in .NET Core 1.0+, .NET 5+, and .NET Standard 1.0+.
            // This information will be written to the project's README file.

            string readmePath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "README.md");

            string content = @"# Project README

## .NET Version Requirement for CancellationToken

The `CancellationToken` struct is part of the `System.Threading` namespace and requires **.NET Framework 4.0** or later.
It is also supported in **.NET Core 1.0+**, **.NET 5+**, and **.NET Standard 1.0+**.

Ensure your project targets one of the above frameworks to use `CancellationToken` with Aspose.Words.";

            // Write the content to README.md (overwrites if it already exists)
            File.WriteAllText(readmePath, content);

            Console.WriteLine($"README file has been created/updated at: {readmePath}");
        }
    }
}
