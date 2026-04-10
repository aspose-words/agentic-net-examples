using System;
using Aspose.Words;
using Aspose.Words.Comparing;

public class RevisionLogger
{
    public static void Main()
    {
        // Create the original document with some content.
        Document docOriginal = new Document();
        DocumentBuilder builderOriginal = new DocumentBuilder(docOriginal);
        builderOriginal.Writeln("This is the original document.");
        builderOriginal.Writeln("Second paragraph.");

        // Create the edited document with intentional differences.
        Document docEdited = new Document();
        DocumentBuilder builderEdited = new DocumentBuilder(docEdited);
        builderEdited.Writeln("This is the edited document."); // changed line
        builderEdited.Writeln("Second paragraph with extra text."); // changed line

        // Ensure both documents have no revisions before comparison.
        if (docOriginal.Revisions.Count == 0 && docEdited.Revisions.Count == 0)
        {
            // Perform comparison – revisions will be added to docOriginal.
            docOriginal.Compare(docEdited, "John Doe", DateTime.Now);
        }

        // Iterate through the revisions collection and log type and affected text.
        foreach (Revision revision in docOriginal.Revisions)
        {
            // ParentNode may be null for style definition changes; guard against it.
            string affectedText = revision.ParentNode != null ? revision.ParentNode.GetText().Trim() : "<no node>";
            Console.WriteLine($"Revision type: {revision.RevisionType}, affected text: \"{affectedText}\"");
        }

        // Save the resulting document that contains the revisions.
        docOriginal.Save("ComparisonResult.docx");
    }
}
