using System;
using System.Text.RegularExpressions;
using Aspose.Words;
using Aspose.Words.Replacing;

class FindReplaceInRuns
{
    static void Main()
    {
        // Load the DOCX document.
        Document doc = new Document("Input.docx");

        // Define the text to find and its replacement.
        string oldText = "placeholder";
        string newText = "actual value";

        // Iterate over all Run nodes in the document.
        foreach (Run run in doc.GetChildNodes(NodeType.Run, true))
        {
            // Replace occurrences of the old text within the run's text.
            if (run.Text.Contains(oldText))
                run.Text = run.Text.Replace(oldText, newText);
        }

        // Save the modified document.
        doc.Save("Output.docx");
    }
}
