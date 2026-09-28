using System;
using System.IO;

public class Program
{
    public static void Main()
    {
        // Path for the README file
        string readmePath = "README.md";

        // Content describing the required .NET version for CancellationToken support
        string content = "# Project README\r\n\r\n" +
                         "## .NET Version Requirement for CancellationToken Support\r\n\r\n" +
                         "- .NET Framework 4.0 and later\r\n" +
                         "- .NET Core 2.0 and later\r\n" +
                         "- .NET 5.0 and later\r\n\r\n" +
                         "The `CancellationToken` struct is available in the `System.Threading` namespace starting from the versions listed above.\r\n";

        // Write the README file
        File.WriteAllText(readmePath, content);

        // Verify that the file was created successfully
        if (!File.Exists(readmePath))
        {
            throw new Exception("Failed to create README file.");
        }
    }
}
