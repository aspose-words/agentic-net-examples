using System;
using System.Text.RegularExpressions;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Add sample paragraphs that contain double (or more) spaces.
        builder.Writeln("This  is  a  sample  text  with  double  spaces.");
        builder.Writeln("Another   line   with   triple   spaces.");

        // Define a regular expression that matches two or more consecutive space characters.
        Regex doubleSpacePattern = new Regex(@" {2,}");

        // Perform the replacement throughout the whole document.
        int replacementCount = doc.Range.Replace(doubleSpacePattern, " ");

        // Validate that at least one replacement was made.
        if (replacementCount == 0)
            throw new InvalidOperationException("No double spaces were found to replace.");

        // Save the modified document to the local file system.
        const string outputPath = "DoubleSpaceReplaced.docx";
        doc.Save(outputPath);

        // Optionally, write a simple confirmation to the console.
        Console.WriteLine($"Replaced {replacementCount} occurrence(s) of double spaces. Output saved to '{outputPath}'.");
    }
}
