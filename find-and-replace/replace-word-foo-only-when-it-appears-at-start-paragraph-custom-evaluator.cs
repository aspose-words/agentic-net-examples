using System;
using System.Text.RegularExpressions;
using Aspose.Words;
using Aspose.Words.Replacing;

namespace ReplaceFooAtParagraphStart
{
    // Custom callback that only replaces matches occurring at the very beginning of a paragraph.
    public class FooStartReplacingCallback : IReplacingCallback
    {
        public ReplaceAction Replacing(ReplacingArgs args)
        {
            // The match offset is the zero‑based position of the match inside the node that contains the start of the match.
            // If the offset is 0, the match begins at the first character of that node.
            // Additionally, we ensure the node is the first child of its paragraph to guarantee it is the start of the paragraph.
            if (args.MatchOffset == 0 && IsFirstNodeInParagraph(args.MatchNode))
            {
                // Replace "foo" with the desired text.
                args.Replacement = "bar";
                return ReplaceAction.Replace;
            }

            // Skip any occurrence that is not at the start of a paragraph.
            return ReplaceAction.Skip;
        }

        // Helper to verify that the node containing the match is the first child of its paragraph.
        private bool IsFirstNodeInParagraph(Node matchNode)
        {
            if (matchNode == null) return false;

            Paragraph paragraph = matchNode.ParentNode as Paragraph;
            if (paragraph == null) return false;

            // The first child of a paragraph is typically a Run (or other inline node).
            return paragraph.FirstChild == matchNode;
        }
    }

    class Program
    {
        static void Main()
        {
            // Load an existing document.
            Document doc = new Document("Input.docx");

            // Set up find/replace options with the custom callback.
            FindReplaceOptions options = new FindReplaceOptions(new FooStartReplacingCallback());

            // Use a regular expression to find the word "foo". The pattern \bfoo\b matches the whole word.
            Regex fooPattern = new Regex(@"\bfoo\b", RegexOptions.IgnoreCase);

            // Perform the replace operation. The callback will decide whether to replace each match.
            doc.Range.Replace(fooPattern, "bar", options);

            // Save the modified document.
            doc.Save("Output.docx");
        }
    }
}
