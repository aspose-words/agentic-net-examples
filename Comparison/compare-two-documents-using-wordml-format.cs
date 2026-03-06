using System;
using Aspose.Words;
using Aspose.Words.Comparing;

class Program
{
    static void Main()
    {
        // Paths to the source documents.
        string originalPath = "Original.docx";
        string editedPath   = "Edited.docx";

        // Load the documents.
        Document docOriginal = new Document(originalPath);
        Document docEdited   = new Document(editedPath);

        // Ensure both documents have no revisions before comparison.
        if (docOriginal.Revisions.Count == 0 && docEdited.Revisions.Count == 0)
        {
            // Compare the edited document against the original.
            // The author name and current date/time are used for the generated revisions.
            docOriginal.Compare(docEdited, "Author", DateTime.Now);
        }

        // Save the comparison result in WORDML (XML) format.
        string resultPath = "ComparisonResult.xml";
        docOriginal.Save(resultPath, SaveFormat.WordML);
    }
}
