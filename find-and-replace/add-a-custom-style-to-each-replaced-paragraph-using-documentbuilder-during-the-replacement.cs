using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Replacing;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Add sample paragraphs that contain the text to be replaced.
        builder.Writeln("Dear _Name_, welcome to our service.");
        builder.Writeln("Your account _Name_ has been activated.");
        builder.Writeln("Please contact _Name_ for further assistance.");

        // Define a custom paragraph style.
        Style customStyle = doc.Styles.Add(StyleType.Paragraph, "MyCustomStyle");
        customStyle.Font.Name = "Arial";
        customStyle.Font.Size = 14;
        customStyle.Font.Color = System.Drawing.Color.Blue; // Use fully‑qualified System.Drawing color.
        customStyle.Font.Bold = true;

        // Set up the find‑replace options with a custom callback.
        FindReplaceOptions options = new FindReplaceOptions(new StyleApplyingCallback(doc, customStyle));

        // Perform the replacement. Each paragraph that contains a match will receive the custom style.
        int replacementCount = doc.Range.Replace("_Name_", "John", options);

        // Validate that at least one replacement occurred.
        if (replacementCount == 0)
            throw new InvalidOperationException("No occurrences of the target text were found.");

        // Save the modified document to the local file system.
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "ReplacedWithStyle.docx");
        doc.Save(outputPath);

        // Indicate success.
        Console.WriteLine($"Replacements performed: {replacementCount}");
        Console.WriteLine($"Document saved to: {outputPath}");
    }

    // Callback that applies the custom style to the paragraph containing each match.
    private class StyleApplyingCallback : IReplacingCallback
    {
        private readonly Document _document;
        private readonly Style _style;

        public StyleApplyingCallback(Document document, Style style)
        {
            _document = document ?? throw new ArgumentNullException(nameof(document));
            _style = style ?? throw new ArgumentNullException(nameof(style));
        }

        ReplaceAction IReplacingCallback.Replacing(ReplacingArgs args)
        {
            // Locate the paragraph that holds the beginning of the match.
            Paragraph paragraph = args.MatchNode.GetAncestor(NodeType.Paragraph) as Paragraph;
            if (paragraph != null)
            {
                // Use DocumentBuilder to move to the paragraph and apply the custom style.
                DocumentBuilder builder = new DocumentBuilder(_document);
                builder.MoveTo(paragraph);
                builder.ParagraphFormat.Style = _style;
            }

            // Replace the matched text with the desired replacement.
            args.Replacement = "John";

            return ReplaceAction.Replace;
        }
    }
}
