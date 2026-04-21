using System;
using Aspose.Words;
using Aspose.Words.Replacing;

namespace AsposeWordsTrackChangesExample
{
    public class Program
    {
        public static void Main()
        {
            // Create the original document.
            Document originalDoc = new Document();
            DocumentBuilder builder = new DocumentBuilder(originalDoc);
            builder.Writeln("This is the original document.");
            builder.Writeln("It contains a single paragraph.");

            // Create the edited document.
            Document editedDoc = new Document();
            builder = new DocumentBuilder(editedDoc);
            builder.Writeln("This is the edited document.");
            builder.Writeln("It contains a modified paragraph.");

            // Ensure both documents have no revisions before comparison.
            if (originalDoc.HasRevisions || editedDoc.HasRevisions)
                throw new InvalidOperationException("Documents must not contain revisions before comparison.");

            // Compare the documents, generating revisions in the original document.
            string authorName = "John Doe";
            DateTime comparisonDate = DateTime.Now;
            originalDoc.Compare(editedDoc, authorName, comparisonDate);

            // Output revision details.
            Console.WriteLine($"Revisions generated after comparison (Author: {authorName}):");
            foreach (Revision revision in originalDoc.Revisions)
            {
                Console.WriteLine($"- Type: {revision.RevisionType}, Author: {revision.Author}, Text: \"{revision.ParentNode.GetText().Trim()}\"");
            }

            // Save the document that now contains revisions.
            string outputPath = "ComparedDocument.docx";
            originalDoc.Save(outputPath);
            Console.WriteLine($"Compared document saved to: {outputPath}");
        }
    }
}
