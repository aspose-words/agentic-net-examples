using System;
using Aspose.Words;

class Program
{
    static void Main()
    {
        // Load the document that contains tracked changes.
        Document doc = new Document("Input.docx"); // replace with actual path

        // Iterate through all revision groups in the document.
        foreach (RevisionGroup group in doc.Revisions.Groups)
        {
            // Identify groups that consist of consecutive delete revisions.
            if (group.RevisionType == RevisionType.Deletion)
            {
                // Accept each revision that belongs to this delete group.
                // After acceptance the deleted text is removed permanently.
                foreach (Revision rev in doc.Revisions)
                {
                    if (rev.Group == group)
                        rev.Accept();
                }
            }
        }

        // Save the document after the merged deletions have been accepted.
        doc.Save("Output.docx"); // replace with desired output path
    }
}
