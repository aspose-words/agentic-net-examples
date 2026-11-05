using System;
using System.IO;
using System.Text;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Create the original document.
        Document originalDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(originalDoc);
        builder.Writeln("Hello world.");
        builder.Writeln("This is a sample document.");

        // Save the original document to a memory stream.
        using (MemoryStream originalStream = new MemoryStream())
        {
            originalDoc.Save(originalStream, SaveFormat.Docx);
            originalStream.Position = 0;

            // Load the original content into a new document that will be modified.
            Document revisedDoc = new Document(originalStream);

            // Make some changes to the revised document.
            DocumentBuilder revBuilder = new DocumentBuilder(revisedDoc);

            // Modify existing text.
            Paragraph firstParagraph = revisedDoc.FirstSection.Body.Paragraphs[0];
            if (firstParagraph.Runs.Count > 0)
                firstParagraph.Runs[0].Text = "Hello Aspose.Words!";

            // Add a new paragraph.
            revBuilder.MoveToDocumentEnd();
            revBuilder.Writeln("Added a new paragraph.");

            // Compare the revised document against the original to generate revisions.
            revisedDoc.Compare(originalDoc, "Comparer", DateTime.Now);

            // Build a revision report in memory.
            StringBuilder report = new StringBuilder();
            foreach (Revision rev in revisedDoc.Revisions)
            {
                string revText = rev.ParentNode?.GetText()?.Trim() ?? string.Empty;
                report.AppendLine(
                    $"Type: {rev.RevisionType}, Author: {rev.Author}, Date: {rev.DateTime}, Text: \"{revText}\"");
            }

            // Output the revision report.
            Console.WriteLine(report.ToString());
        }
    }
}
