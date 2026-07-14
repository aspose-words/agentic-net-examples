using System;
using System.Text.RegularExpressions;
using Aspose.Words;
using Aspose.Words.Replacing;

public class Program
{
    public static void Main()
    {
        // Create a new document and add some paragraphs.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("This paragraph will stay unchanged.");
        builder.Writeln("Please replace this text.");
        builder.Writeln("Another line that needs to be replace.");

        // Define a custom paragraph style that will be applied to replaced paragraphs.
        const string styleName = "ReplacedParagraphStyle";
        Style customStyle = doc.Styles.Add(StyleType.Paragraph, styleName);
        customStyle.Font.Name = "Arial";
        customStyle.Font.Size = 14;
        customStyle.Font.Color = System.Drawing.Color.Blue;
        customStyle.Font.Bold = true;

        // Set up the find‑replace options with a callback that applies the style.
        FindReplaceOptions options = new FindReplaceOptions
        {
            ReplacingCallback = new ReplaceAndStyleCallback(doc, styleName)
        };

        // Perform the replacement.
        int replacedCount = doc.Range.Replace("replace", "replaced", options);
        if (replacedCount == 0)
            throw new InvalidOperationException("No occurrences were replaced.");

        // Save the modified document.
        doc.Save("output.docx");
        Console.WriteLine($"Replacements performed: {replacedCount}");
    }

    // Callback that applies the custom style to the paragraph containing each match.
    private class ReplaceAndStyleCallback : IReplacingCallback
    {
        private readonly Document _document;
        private readonly string _styleName;

        public ReplaceAndStyleCallback(Document document, string styleName)
        {
            _document = document ?? throw new ArgumentNullException(nameof(document));
            _styleName = styleName ?? throw new ArgumentNullException(nameof(styleName));
        }

        public ReplaceAction Replacing(ReplacingArgs args)
        {
            // The match is inside a Run; its parent is the Paragraph we want to style.
            if (args.MatchNode?.ParentNode is Paragraph paragraph)
            {
                paragraph.ParagraphFormat.Style = _document.Styles[_styleName];
            }

            // Continue with the normal replacement.
            return ReplaceAction.Replace;
        }
    }
}
