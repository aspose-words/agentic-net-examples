using System;
using System.IO;
using System.Text.RegularExpressions;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Path for the output document (saved in the program's working directory).
        string outputPath = "Output.docx";

        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Add sample paragraphs containing multiple consecutive spaces.
        builder.Writeln("This  is   a    sample     text with  multiple   spaces.");
        builder.Writeln("Another    line   with    spaces.");
        builder.Writeln("   Leading and trailing   spaces   should also be handled.   ");

        // Define a regular expression that matches two or more space characters.
        Regex multipleSpaces = new Regex(@" {2,}");

        // Replace all occurrences of the pattern with a single space.
        int replacementCount = doc.Range.Replace(multipleSpaces, " ");

        // Ensure that at least one replacement was performed.
        if (replacementCount == 0)
        {
            throw new InvalidOperationException("No multiple spaces were found to replace.");
        }

        // Save the modified document.
        doc.Save(outputPath);

        // Optional: write a short confirmation to the console.
        Console.WriteLine($"Replaced {replacementCount} occurrences of multiple spaces. Document saved to '{Path.GetFullPath(outputPath)}'.");
    }
}
