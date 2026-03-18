using System;
using Aspose.Words;

class Program
{
    static void Main()
    {
        // Load the document that contains tracked changes.
        Document doc1 = new Document("input.docx");

        // Iterate through each revision in the document.
        foreach (Revision revision in doc1.Revisions)
        {
            // Determine the affected text.
            string affectedText;

            // StyleDefinitionChange revisions affect a style, not a node.
            if (revision.RevisionType == RevisionType.StyleDefinitionChange && revision.ParentStyle != null)
                affectedText = $"Style: {revision.ParentStyle.Name}";
            else
                affectedText = revision.ParentNode?.GetText()?.Trim() ?? string.Empty;

            // Log the revision type and the affected text.
            Console.WriteLine($"Revision type: {revision.RevisionType}, text: [{affectedText}]");
        }

        // Save the document (optional, if you need to persist any changes).
        doc1.Save("output.docx");
    }
}
