using System;
using System.Text;
using System.Text.RegularExpressions;
using Aspose.Words;
using Aspose.Words.Replacing;

namespace AsposeWordsFindReplaceDemo
{
    // Callback that replaces a figure caption and inserts a Table of Figures after it.
    public class FigureCaptionCallback : IReplacingCallback
    {
        public ReplaceAction Replacing(ReplacingArgs args)
        {
            // Replace "Figure" with "Fig." in the matched caption.
            args.Replacement = args.Match.Value.Replace("Figure", "Fig.");

            // Insert a Table of Figures after the paragraph that contains the matched caption.
            if (args.MatchNode?.ParentNode is Paragraph paragraph)
            {
                // DocumentBuilder works with Document, so cast the base document to Document.
                var builder = new DocumentBuilder((Document)args.MatchNode.Document);
                builder.MoveTo(paragraph);
                builder.Writeln(); // Ensure a new paragraph after the caption.

                // Insert a Table of Figures field. The switch "\h" makes entries hyperlinked,
                // and "\f \"Figure\"" tells Word to use the "Figure" style as the caption identifier.
                builder.InsertTableOfContents("\\h \\f \"Figure\"");
            }

            // Proceed with the normal replacement.
            return ReplaceAction.Replace;
        }
    }

    public class Program
    {
        public static void Main()
        {
            // Create a new blank document and add sample content with figure captions.
            var doc = new Document();
            var builder = new DocumentBuilder(doc);

            builder.Writeln("Introduction text.");
            builder.Writeln("Figure 1: First sample figure.");
            builder.Writeln("Some explanatory text after the first figure.");
            builder.Writeln("Figure 2: Second sample figure.");
            builder.Writeln("Conclusion text.");

            // Save the original document (optional, for inspection).
            doc.Save("original.docx");

            // Set up the find-and-replace operation with a custom callback.
            var options = new FindReplaceOptions
            {
                ReplacingCallback = new FigureCaptionCallback()
            };

            // Regex that matches an entire figure caption line.
            var figureCaptionRegex = new Regex(@"Figure \d+:.*", RegexOptions.Multiline);

            // Perform the replace. The callback will modify the text and insert a Table of Figures.
            int replacedCount = doc.Range.Replace(figureCaptionRegex, "$0", options);

            // Validate that at least one caption was processed.
            if (replacedCount == 0)
                throw new InvalidOperationException("No figure captions were found for replacement.");

            // Update fields so that the inserted Table of Figures reflects the current captions.
            doc.UpdateFields();

            // Save the modified document.
            doc.Save("output.docx");
        }
    }
}
