using System;
using Aspose.Words;
using Aspose.Words.Comparing;

class CompareDocuments
{
    static void Main()
    {
        // Load the original and edited DOCX files.
        Document docOriginal = new Document("Original.docx");
        Document docEdited = new Document("Edited.docx");

        // Documents must not contain revisions before comparison.
        if (docOriginal.Revisions.Count != 0 || docEdited.Revisions.Count != 0)
            throw new InvalidOperationException("Both documents must be revision‑free before comparison.");

        // Configure comparison options (optional).
        CompareOptions options = new CompareOptions
        {
            Granularity = Granularity.WordLevel,   // Track changes by word.
            IgnoreFormatting = true,               // Ignore formatting differences.
            Target = ComparisonTargetType.New      // Use the edited document as the base.
        };

        // Compare the documents; revisions are added to docOriginal.
        docOriginal.Compare(docEdited, "JD", DateTime.Now, options);

        // List all revisions created by the comparison.
        foreach (Revision rev in docOriginal.Revisions)
        {
            Console.WriteLine($"Revision type: {rev.RevisionType}, Node type: {rev.ParentNode.NodeType}");
            Console.WriteLine($"\tChanged text: \"{rev.ParentNode.GetText().Trim()}\"");
        }

        // Accept all revisions to make docOriginal identical to docEdited.
        docOriginal.Revisions.AcceptAll();

        // Save the resulting document.
        docOriginal.Save("Result.docx");
    }
}
