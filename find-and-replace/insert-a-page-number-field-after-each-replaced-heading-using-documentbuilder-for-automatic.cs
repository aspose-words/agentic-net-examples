using System;
using Aspose.Words;
using Aspose.Words.Replacing;
using Aspose.Words.Fields;
using Aspose.Words.Drawing;

namespace AsposeWordsFindReplaceDemo
{
    // Callback that inserts a PAGE field after each replaced heading.
    public class InsertPageNumberAfterReplace : IReplacingCallback
    {
        private readonly Document _doc;

        public InsertPageNumberAfterReplace(Document doc)
        {
            _doc = doc ?? throw new ArgumentNullException(nameof(doc));
        }

        public ReplaceAction Replacing(ReplacingArgs args)
        {
            // The node that contains the start of the match.
            var matchNode = args.MatchNode;
            if (matchNode?.ParentNode is not Paragraph paragraph)
                return ReplaceAction.Skip;

            // Insert a new paragraph after the matched paragraph.
            var newParagraph = new Paragraph(_doc);
            paragraph.ParentNode.InsertAfter(newParagraph, paragraph);

            // Move a builder to the new paragraph and insert a PAGE field.
            var builder = new DocumentBuilder(_doc);
            builder.MoveTo(newParagraph);
            builder.InsertField(FieldType.FieldPage, true);

            // Allow the original replacement to proceed.
            return ReplaceAction.Replace;
        }
    }

    public class Program
    {
        public static void Main()
        {
            // 1. Create a sample document with headings.
            var doc = new Document();
            var builder = new DocumentBuilder(doc);

            builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Heading1;
            builder.Writeln("Heading One");
            builder.Writeln("Heading Two");
            builder.Writeln("Heading Three");

            // Save the source document.
            const string inputPath = "input.docx";
            doc.Save(inputPath);

            // 2. Load the document for processing.
            var loadedDoc = new Document(inputPath);

            // 3. Set up find-and-replace with a callback to insert page numbers.
            var options = new FindReplaceOptions
            {
                ReplacingCallback = new InsertPageNumberAfterReplace(loadedDoc)
            };

            // Replace the word "Heading" with "Section".
            int replacedCount = loadedDoc.Range.Replace("Heading", "Section", options);
            if (replacedCount == 0)
                throw new InvalidOperationException("Expected at least one replacement.");

            // 4. Update fields so PAGE fields show correct numbers.
            loadedDoc.UpdateFields();

            // 5. Save the modified document.
            const string outputPath = "output.docx";
            loadedDoc.Save(outputPath);
        }
    }
}
