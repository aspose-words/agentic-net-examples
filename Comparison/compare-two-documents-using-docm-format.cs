using System;
using Aspose.Words;
using Aspose.Words.Comparing;

class CompareDocm
{
    static void Main()
    {
        // Load the original DOCM document.
        Document original = new Document("Original.docm");

        // Load the edited DOCM document.
        Document edited = new Document("Edited.docm");

        // Ensure both documents have no revisions before performing the comparison.
        if (original.Revisions.Count == 0 && edited.Revisions.Count == 0)
        {
            // Compare the documents. The revisions will be added to the original document.
            // "JD" is the author initials for the revisions, and the current time is used as the revision timestamp.
            original.Compare(edited, "JD", DateTime.Now);
        }

        // Save the result of the comparison as a DOCM file.
        original.Save("Compared.docm");
    }
}
