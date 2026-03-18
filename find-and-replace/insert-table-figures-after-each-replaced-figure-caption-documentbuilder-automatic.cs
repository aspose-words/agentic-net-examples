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
            // captionParagraph.Document returns DocumentBase, so cast to Document which the constructor expects.
            DocumentBuilder builder = new DocumentBuilder((Document)captionParagraph.Document);

            // Move the cursor to the end of the caption paragraph.
            // MoveTo places the cursor at the start of the node, so we first move to the paragraph,
            // then insert a paragraph break to position the cursor after it.
            builder.MoveTo(captionParagraph);
            builder.InsertBreak(BreakType.ParagraphBreak);

            // Insert a TOC field that will act as a Table of Figures.
            // FieldType.FieldTOC creates a generic TOC field; we configure it to build a Table of Figures.
            FieldToc tocField = (FieldToc)builder.InsertField(FieldType.FieldTOC, true);
            tocField.TableOfFiguresLabel = "Figure";          // Use the "Figure" sequence identifier.
            tocField.InsertHyperlinks = true;                // Make entries clickable.
            tocField.CaptionlessTableOfFiguresLabel = "";    // Not needed for standard TOF.

            // Skip the original match text (we already removed it by replacing with an empty string).
            return ReplaceAction.Skip;
        }
    }

    class Program
    {
        static void Main()
        {
            // Load the source document that contains figure captions.
            Document doc = new Document("InputDocument.docx");

            // Set up find-and-replace options with our custom callback.
            FindReplaceOptions options = new FindReplaceOptions();
            options.ReplacingCallback = new InsertTableOfFiguresAfterCaption();

            // Regular expression that matches typical figure captions, e.g., "Figure 1: Description".
            // Adjust the pattern as needed for your document's caption format.
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
