using System;
using Aspose.Words;
using Aspose.Words.Replacing;

public class Program
{
    public static void Main()
    {
        // Create a sample document with placeholders.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("First paragraph with placeholder.");
        builder.Writeln("Second paragraph without.");
        builder.Writeln("Third paragraph with placeholder.");

        // Save the sample document locally.
        const string inputPath = "input.docx";
        doc.Save(inputPath);

        // Load the document for processing.
        Document loaded = new Document(inputPath);

        // Define a custom paragraph style.
        const string customStyleName = "MyCustomStyle";
        Style customStyle = loaded.Styles.Add(StyleType.Paragraph, customStyleName);
        // Font color is omitted to avoid System.Drawing usage per category rules.
        customStyle.Font.Size = 16;
        customStyle.Font.Bold = true;

        // Set up find‑replace with a callback that applies the custom style.
        FindReplaceOptions options = new FindReplaceOptions
        {
            ReplacingCallback = new ReplaceAndStyleCallback(customStyleName)
        };

        int replacedCount = loaded.Range.Replace("placeholder", "REPLACED", options);
        if (replacedCount == 0)
            throw new InvalidOperationException("Expected at least one replacement.");

        // Save the modified document.
        const string outputPath = "output.docx";
        loaded.Save(outputPath);
    }

    // Callback that replaces the matched text and styles the containing paragraph.
    private class ReplaceAndStyleCallback : IReplacingCallback
    {
        private readonly string _styleName;

        public ReplaceAndStyleCallback(string styleName)
        {
            _styleName = styleName;
        }

        public ReplaceAction Replacing(ReplacingArgs args)
        {
            // Perform the text replacement.
            args.Replacement = "REPLACED";

            // Apply the custom style to the paragraph that contains the match.
            if (args.MatchNode?.ParentNode is Paragraph paragraph)
            {
                DocumentBuilder builder = new DocumentBuilder((Document)paragraph.Document);
                builder.MoveTo(paragraph);
                builder.ParagraphFormat.StyleName = _styleName;
            }

            return ReplaceAction.Replace;
        }
    }
}
