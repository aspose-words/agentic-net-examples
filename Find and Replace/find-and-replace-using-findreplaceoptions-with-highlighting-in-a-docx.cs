using System;
using System.Drawing;
using System.Text.RegularExpressions;
using Aspose.Words;
using Aspose.Words.Replacing;

namespace FindReplaceHighlightExample
{
    class Program
    {
        static void Main()
        {
            // Create a new blank document.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Add some sample text that contains the word we want to replace.
            builder.Writeln("Ruby bought a ruby necklace. The ruby was very shiny.");

            // Configure find/replace options.
            FindReplaceOptions options = new FindReplaceOptions();

            // Highlight the replaced text with a light gray background.
            options.ApplyFont.HighlightColor = Color.LightGray;

            // Perform the replace operation.
            // The pattern "ruby" will be replaced with "jade".
            // The replacement will inherit the highlighting defined above.
            doc.Range.Replace("ruby", "jade", options);

            // Save the resulting document.
            doc.Save("FindReplaceHighlight.docx");
        }
    }
}
