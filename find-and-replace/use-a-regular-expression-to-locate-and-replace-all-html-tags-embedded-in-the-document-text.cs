using System;
using System.IO;
using System.Text.RegularExpressions;
using Aspose.Words;
using Aspose.Words.Replacing;
using Aspose.Drawing;          // Required package, not used directly
using Newtonsoft.Json;        // Required package, not used directly

namespace AsposeWordsFindReplaceHtmlTags
{
    public class Program
    {
        public static void Main()
        {
            // Define file names for the sample input and output documents.
            const string inputPath = "input.docx";
            const string outputPath = "output.docx";

            // -----------------------------------------------------------------
            // Create a sample document containing HTML tags embedded in the text.
            // -----------------------------------------------------------------
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Sample text with various HTML tags.
            builder.Writeln("This is a paragraph with <b>bold</b> text.");
            builder.Writeln("Here is a link: <a href=\"https://example.com\">Example</a>.");
            builder.Writeln("An image tag: <img src=\"image.png\" alt=\"Sample Image\"/>");
            builder.Writeln("End of <span style=\"color:red;\">sample</span> document.");

            // Save the created document so that we have a tangible source file.
            doc.Save(inputPath);

            // ---------------------------------------------------------------
            // Load the document from the file system (demonstrates load workflow).
            // ---------------------------------------------------------------
            Document loadedDoc = new Document(inputPath);

            // ---------------------------------------------------------------
            // Define a regular expression that matches any HTML tag.
            // The pattern <[^>]+> captures a '<', then any characters except '>', then a '>'.
            // ---------------------------------------------------------------
            Regex htmlTagRegex = new Regex(@"<[^>]+>", RegexOptions.IgnoreCase);

            // Perform the replacement: remove all HTML tags.
            FindReplaceOptions replaceOptions = new FindReplaceOptions();
            int replacementCount = loadedDoc.Range.Replace(htmlTagRegex, string.Empty, replaceOptions);

            // Validate that at least one replacement occurred.
            if (replacementCount == 0)
                throw new InvalidOperationException("No HTML tags were found to replace.");

            // ---------------------------------------------------------------
            // Save the modified document.
            // ---------------------------------------------------------------
            loadedDoc.Save(outputPath);

            // Optional: write a short confirmation to the console.
            Console.WriteLine($"Replaced {replacementCount} HTML tag(s). Output saved to '{outputPath}'.");
        }
    }
}
