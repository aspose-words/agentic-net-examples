using System;
using System.Text.RegularExpressions;
using Aspose.Words;
using Aspose.Words.Replacing;
using Aspose.Words.Fields;

namespace AsposeWordsTableOfFiguresExample
{
    // Callback that is invoked for each figure caption match.
    // Inserts a Table of Figures (TOC field) right after the paragraph that contains the caption.
    class InsertTableOfFiguresAfterCaption : IReplacingCallback
    {
        public ReplaceAction Replacing(ReplacingArgs args)
        {
            // The match node is a Run inside the paragraph that holds the caption.
            // Get the parent paragraph.
            Paragraph captionParagraph = (Paragraph)args.MatchNode.ParentNode;

            // Create a builder attached to the same document.
            DocumentBuilder builder = new DocumentBuilder((Document)captionParagraph.Document);

            // Move the cursor to the end of the caption paragraph.
            builder.MoveTo(captionParagraph);
            builder.InsertBreak(BreakType.ParagraphBreak);

            // Insert a TOC field that will act as a Table of Figures.
            FieldToc tocField = (FieldToc)builder.InsertField(FieldType.FieldTOC, true);
            tocField.TableOfFiguresLabel = "Figure";
            tocField.InsertHyperlinks = true;
            tocField.CaptionlessTableOfFiguresLabel = string.Empty;

            // Skip the original match text (we already removed it by replacing with an empty string).
            return ReplaceAction.Skip;
        }
    }

    class Program
    {
        static void Main()
        {
            // Create a sample document with a figure caption.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);
            builder.Writeln("This is an introductory paragraph.");
            builder.Writeln("Figure 1: Sample figure caption.");
            builder.Writeln("Some more text after the figure.");

            // Set up find-and-replace options with our custom callback.
            FindReplaceOptions options = new FindReplaceOptions
            {
                ReplacingCallback = new InsertTableOfFiguresAfterCaption()
            };

            // Regular expression that matches typical figure captions, e.g., "Figure 1: Description".
            Regex figureCaptionPattern = new Regex(@"Figure\s+\d+\s*:", RegexOptions.IgnoreCase);

            // Replace each caption with an empty string (the callback will insert the TOF after it).
            doc.Range.Replace(figureCaptionPattern, string.Empty, options);

            // Update all fields so the newly inserted Table of Figures reflects the current content.
            doc.UpdateFields();

            // Save the modified document.
            doc.Save("OutputDocument.docx");
        }
    }
}
