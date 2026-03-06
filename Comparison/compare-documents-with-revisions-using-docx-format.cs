using System;
using Aspose.Words;
using Aspose.Words.Comparing;

class CompareDocuments
{
    static void Main()
    {
        // Create the original document and add a paragraph.
        Document docOriginal = new Document();
        DocumentBuilder builder = new DocumentBuilder(docOriginal);
        builder.Writeln("This is the original document.");

        // Save the original document (optional, demonstrates the save rule).
        docOriginal.Save("Original.docx");

        // Clone the original document to create an edited version.
        Document docEdited = (Document)docOriginal.Clone(true);

        // Modify the cloned document's first paragraph.
        Paragraph firstParagraph = docEdited.FirstSection.Body.FirstParagraph;
        firstParagraph.Runs[0].Text = "This is the edited document.";

        // Save the edited document (optional).
        docEdited.Save("Edited.docx");

        // Ensure both documents have no revisions before comparison.
        if (docOriginal.Revisions.Count == 0 && docEdited.Revisions.Count == 0)
        {
            // Compare the documents. Revisions will be added to docOriginal.
            docOriginal.Compare(docEdited, "JD", DateTime.Now);
        }

        // Output details of each revision created by the comparison.
        foreach (Revision rev in docOriginal.Revisions)
        {
            Console.WriteLine($"Revision type: {rev.RevisionType}, Node type: {rev.ParentNode.NodeType}");
            Console.WriteLine($"\tChanged text: \"{rev.ParentNode.GetText().Trim()}\"");
        }

        // Accept all revisions so that docOriginal becomes identical to docEdited.
        docOriginal.AcceptAllRevisions();

        // Verify that the texts of both documents are now equal.
        bool textsEqual = docOriginal.GetText() == docEdited.GetText();
        Console.WriteLine($"Documents equal after accepting revisions: {textsEqual}");

        // Save the final document after accepting revisions.
        docOriginal.Save("Result.docx");
    }
}
