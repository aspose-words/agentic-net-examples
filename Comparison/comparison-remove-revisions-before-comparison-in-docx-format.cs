using System;
using Aspose.Words;
using Aspose.Words.Comparing;

class RemoveRevisionsBeforeComparison
{
    static void Main()
    {
        // Path to the folder that contains the documents.
        string dataDir = @"C:\Data\";

        // Load the original and the edited documents.
        Document docOriginal = new Document(dataDir + "Original.docx");
        Document docEdited = new Document(dataDir + "Edited.docx");

        // Accept (remove) all revisions in both documents so that the comparison works on clean content.
        docOriginal.Revisions.AcceptAll();
        docEdited.Revisions.AcceptAll();

        // Set up comparison options (default options are sufficient for a basic comparison).
        CompareOptions compareOptions = new CompareOptions();

        // Perform the comparison. The result (revisions that represent differences) will be stored in docOriginal.
        docOriginal.Compare(docEdited, "Comparer", DateTime.Now, compareOptions);

        // Save the comparison result to a new DOCX file.
        docOriginal.Save(dataDir + "ComparisonResult.docx");
    }
}
