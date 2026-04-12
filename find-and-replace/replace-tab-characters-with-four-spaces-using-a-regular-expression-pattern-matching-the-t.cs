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

        // Insert sample text that contains tab characters.
        builder.Writeln("Name\tAge\tLocation");
        builder.Writeln("Alice\t30\tNew York");
        builder.Writeln("Bob\t25\tLondon");

        // Define a regular expression that matches a tab character.
        Regex tabRegex = new Regex(@"\t");

        // Replace each tab with four spaces.
        int replacementCount = doc.Range.Replace(tabRegex, "    ");

        // Ensure that at least one replacement was made.
        if (replacementCount == 0)
            throw new InvalidOperationException("No tab characters were found to replace.");

        // Save the modified document to the local file system.
        doc.Save("Modified.docx");
    }
}
