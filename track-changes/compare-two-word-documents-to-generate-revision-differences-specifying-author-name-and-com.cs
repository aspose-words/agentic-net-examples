using System;
using Aspose.Words;
using Aspose.Words.Comparing;

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
            builder.Writeln("It has two paragraphs.");

            // Create the edited document.
            Document editedDoc = new Document();
            builder = new DocumentBuilder(editedDoc);
            builder.Writeln("This is the edited document."); // Changed text.
            builder.Writeln("It has three paragraphs now."); // Modified paragraph.
            builder.Writeln("Additional paragraph added.");   // New paragraph.

            // Ensure both documents have no revisions before comparison.
            if (originalDoc.HasRevisions || editedDoc.HasRevisions)
                throw new InvalidOperationException("Documents must not contain revisions before comparison.");

            // Compare the documents. The original document will receive revisions describing the differences.
            string authorName = "Comparer";
            DateTime comparisonDate = DateTime.Now;
            originalDoc.Compare(editedDoc, authorName, comparisonDate);

            // Output revision details to the console.
            Console.WriteLine("Revisions found after comparison:");
            foreach (Revision revision in originalDoc.Revisions)
            {
                Console.WriteLine($"Author: {revision.Author}");
                Console.WriteLine($"Date: {revision.DateTime}");
                Console.WriteLine($"Type: {revision.RevisionType}");
                Console.WriteLine($"Changed text: \"{revision.ParentNode.GetText().Trim()}\"");
                Console.WriteLine(new string('-', 40));
            }

            // Save the document that now contains the revisions.
            originalDoc.Save("ComparedDocument.docx");
        }
    }
}
