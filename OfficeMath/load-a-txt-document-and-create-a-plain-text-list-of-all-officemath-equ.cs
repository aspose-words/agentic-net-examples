using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Loading;
using Aspose.Words.Saving;

class OfficeMathExtractor
{
    static void Main()
    {
        // Path to the input TXT file that contains OfficeMath equations exported as plain text.
        string txtFilePath = @"C:\Docs\Equations.txt";

        // Load the TXT file using PlainTextDocument which automatically detects the format.
        // No additional load options are required for plain‑text files.
        PlainTextDocument plainDoc = new PlainTextDocument(txtFilePath);

        // Retrieve the full text content of the document.
        string fullText = plainDoc.Text;

        // Split the text into separate lines, removing empty entries.
        // Each line is assumed to represent a single OfficeMath equation string.
        string[] equationLines = fullText.Split(
            new[] { '\r', '\n' },
            StringSplitOptions.RemoveEmptyEntries);

        // Output the list of extracted equation strings.
        Console.WriteLine("Extracted OfficeMath equations:");
        foreach (string equation in equationLines)
        {
            Console.WriteLine("- " + equation.Trim());
        }
    }
}
