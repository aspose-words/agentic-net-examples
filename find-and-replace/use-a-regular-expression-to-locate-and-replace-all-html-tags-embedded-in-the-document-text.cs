using System;
using System.Text.RegularExpressions;
using Aspose.Words;
using Aspose.Words.Replacing;

namespace AsposeWordsFindReplaceExample
{
    public class Program
    {
        public static void Main()
        {
            // Create a new blank document.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Insert sample text that contains HTML tags.
            builder.Writeln("This is a sample paragraph with <b>bold</b> text,");
            builder.Writeln("an <a href='https://example.com'>example link</a>,");
            builder.Writeln("and an <img src='image.png' alt='image'/> tag.");

            // Save the original document (optional, for inspection).
            doc.Save("input.docx");

            // Define a regular expression that matches any HTML tag.
            Regex htmlTagRegex = new Regex(@"<[^>]+>", RegexOptions.Compiled);

            // Perform the replacement: remove all HTML tags.
            FindReplaceOptions options = new FindReplaceOptions();
            int replacedCount = doc.Range.Replace(htmlTagRegex, string.Empty, options);

            // Ensure that at least one replacement was made.
            if (replacedCount == 0)
                throw new InvalidOperationException("No HTML tags were found to replace.");

            // Save the modified document.
            doc.Save("output.docx");

            // Write a simple confirmation to the console.
            Console.WriteLine($"Replaced {replacedCount} HTML tag(s). Output saved to 'output.docx'.");
        }
    }
}
