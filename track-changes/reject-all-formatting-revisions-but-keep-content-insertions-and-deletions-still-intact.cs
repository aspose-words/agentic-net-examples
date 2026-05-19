using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Replacing;

public class Program
{
    public static void Main()
    {
        // Create the original document.
        Document original = new Document();
        DocumentBuilder builder = new DocumentBuilder(original);
        builder.Writeln("This is a sample paragraph.");
        builder.Writeln("Another line of text.");
        // Save the original to a memory stream for later comparison.
        using (MemoryStream originalStream = new MemoryStream())
        {
            original.Save(originalStream, SaveFormat.Docx);
            originalStream.Position = 0;

            // Load the original document from the stream.
            Document originalDoc = new Document(originalStream);

            // Create an edited version of the document.
            Document editedDoc = (Document)originalDoc.Clone(true);
            DocumentBuilder editBuilder = new DocumentBuilder(editedDoc);

            // ----- Insert a new paragraph (insertion revision) -----
            editBuilder.MoveToDocumentEnd();
            editBuilder.Writeln("Inserted paragraph.");

            // ----- Delete a paragraph (deletion revision) -----
            // Remove the first paragraph.
            Paragraph firstParagraph = editedDoc.FirstSection.Body.Paragraphs[0];
            firstParagraph.Remove();

            // ----- Change formatting (format revision) -----
            // Make the remaining text bold.
            foreach (Paragraph para in editedDoc.FirstSection.Body.Paragraphs)
            {
                foreach (Run run in para.Runs)
                {
                    run.Font.Bold = true;
                }
            }

            // Compare the original with the edited document to generate revisions.
            originalDoc.Compare(editedDoc, "Comparer", DateTime.Now);

            // Reject only formatting revisions, keep insertions and deletions.
            for (int i = originalDoc.Revisions.Count - 1; i >= 0; i--)
            {
                Revision rev = originalDoc.Revisions[i];
                if (rev.RevisionType == RevisionType.FormatChange)
                {
                    rev.Reject();
                }
            }

            // Save the resulting document.
            string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "Result.docx");
            originalDoc.Save(outputPath);
        }
    }
}
