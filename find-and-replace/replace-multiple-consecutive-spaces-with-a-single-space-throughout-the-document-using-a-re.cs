using System;
using System.IO;
using System.Text.RegularExpressions;
using Aspose.Words;
using Aspose.Words.Replacing;

public class Program
{
    public static void Main()
    {
        // Define file names in the current directory.
        string inputPath = Path.Combine(Directory.GetCurrentDirectory(), "input.docx");
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "output.docx");

        // -----------------------------------------------------------------
        // Create a sample document with irregular spacing.
        // -----------------------------------------------------------------
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("This  is   a    sentence   with  multiple   spaces.");
        builder.Writeln("Another    line    with   extra   spaces.");
        builder.Writeln("End   of   document.");
        doc.Save(inputPath);

        // -----------------------------------------------------------------
        // Load the document and replace consecutive spaces with a single space.
        // -----------------------------------------------------------------
        Document loaded = new Document(inputPath);

        // Regular expression that matches two or more space characters.
        Regex multipleSpaces = new Regex(@" {2,}");

        // Perform the replacement across the whole document.
        int replacementCount = loaded.Range.Replace(multipleSpaces, " ", new FindReplaceOptions());

        // Validate that at least one replacement occurred.
        if (replacementCount == 0)
            throw new InvalidOperationException("Expected at least one space reduction replacement.");

        // Save the modified document.
        loaded.Save(outputPath);
    }
}
