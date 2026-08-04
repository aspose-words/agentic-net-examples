using System;
using Aspose.Words;
using Aspose.Words.Replacing;
using Aspose.Words.Fields;

namespace AsposeWordsFindReplaceDemo
{
    // Callback that replaces heading text and inserts a PAGE field after each replaced heading.
    public class HeadingReplaceCallback : IReplacingCallback
    {
        public ReplaceAction Replacing(ReplacingArgs args)
        {
            // Replace the matched heading text with the new text.
            args.Replacement = "Section";

            // Insert a PAGE field after the paragraph that contains the match.
            if (args.MatchNode?.ParentNode is Paragraph paragraph)
            {
                // The Document property of a Paragraph is of type DocumentBase,
                // so we need to cast it to Document before creating a DocumentBuilder.
                var doc = (Document)paragraph.Document;
                var builder = new DocumentBuilder(doc);

                // Move the builder to the paragraph that contains the match.
                builder.MoveTo(paragraph);

                // Insert a new empty paragraph after the current one.
                builder.InsertParagraph();

                // Insert a PAGE field that will display the current page number.
                builder.InsertField(FieldType.FieldPage, true);
            }

            return ReplaceAction.Replace;
        }
    }

    public class Program
    {
        public static void Main()
        {
            // -----------------------------------------------------------------
            // 1. Create a sample document with a few headings.
            // -----------------------------------------------------------------
            var doc = new Document();
            var builder = new DocumentBuilder(doc);

            // Add three headings.
            builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Heading1;
            builder.Writeln("Heading One");
            builder.Writeln("Heading Two");
            builder.Writeln("Heading Three");

            // Save the source document.
            const string inputPath = "input.docx";
            doc.Save(inputPath);

            // -----------------------------------------------------------------
            // 2. Load the document and perform find-and-replace with a callback.
            // -----------------------------------------------------------------
            var loadedDoc = new Document(inputPath);

            var replaceCallback = new HeadingReplaceCallback();
            var options = new FindReplaceOptions
            {
                ReplacingCallback = replaceCallback
            };

            // Replace the word "Heading" with "Section" and trigger the callback.
            int replacedCount = loadedDoc.Range.Replace("Heading", "Section", options);
            if (replacedCount == 0)
                throw new InvalidOperationException("Expected at least one replacement.");

            // Update fields so that PAGE fields show correct numbers.
            loadedDoc.UpdateFields();

            // -----------------------------------------------------------------
            // 3. Save the modified document.
            // -----------------------------------------------------------------
            const string outputPath = "output.docx";
            loadedDoc.Save(outputPath);
        }
    }
}
