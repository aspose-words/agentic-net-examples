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

        // Define the original bullet character (black circle, U+2022).
        const string originalBullet = "\u2022";

        // Add a few paragraphs that start with the original bullet character.
        builder.Writeln($"{originalBullet} First item");
        builder.Writeln($"{originalBullet} Second item");
        builder.Writeln($"{originalBullet} Third item");

        // Save the original document (optional, just to show the before state).
        doc.Save("Original.docx");

        // Prepare a regular expression that matches the original bullet character.
        Regex bulletRegex = new Regex(Regex.Escape(originalBullet));

        // Define the replacement bullet character (white bullet, U+25E6).
        const string newBullet = "\u25E6";

        // Perform the replacement throughout the document.
        int replacementCount = doc.Range.Replace(bulletRegex, newBullet);

        // Ensure that at least one replacement occurred.
        if (replacementCount == 0)
        {
            throw new InvalidOperationException("No bullet characters were replaced.");
        }

        // Save the modified document.
        doc.Save("Modified.docx");
    }
}
