using System;
using Aspose.Words;

class ReplaceInParagraph
{
    static void Main()
    {
        // Load an existing document.
        Document doc = new Document("Input.docx");

        // Get the first paragraph (replace with the desired paragraph index as needed).
        Paragraph paragraph = doc.FirstSection.Body.Paragraphs[0];

        // Text to find and its replacement.
        string findText = "_FullName_";
        string replaceText = "John Doe";

        // Iterate through all runs in the paragraph.
        foreach (Run run in paragraph.Runs)
        {
            // Replace the text inside the run while keeping its original formatting.
            if (run.Text.Contains(findText))
                run.Text = run.Text.Replace(findText, replaceText);
        }

        // Save the modified document.
        doc.Save("Output.docx");
    }
}
