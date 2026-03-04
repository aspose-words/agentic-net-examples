using System;
using Aspose.Words;
using Aspose.Words.Replacing;

class FindAndReplaceExample
{
    static void Main()
    {
        // Load an existing DOCX file.
        Document doc = new Document("Input.docx");

        // DocumentBuilder is required for this task (e.g., to add content before/after replacement if needed).
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Simple find-and-replace: replace placeholder "_FullName_" with "John Doe".
        int replacedCount = doc.Range.Replace("_FullName_", "John Doe");

        // Optionally, use FindReplaceOptions to control the replace behavior (case‑sensitive, whole word, etc.).
        FindReplaceOptions options = new FindReplaceOptions
        {
            MatchCase = false,               // ignore case
            FindWholeWordsOnly = true        // replace only whole words
        };
        // Example of a second replacement using the options.
        doc.Range.Replace("Company", "Acme Corp", options);

        // Save the modified document.
        doc.Save("Output.docx");
    }
}
