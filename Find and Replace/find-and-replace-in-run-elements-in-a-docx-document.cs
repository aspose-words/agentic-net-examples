using System;
using Aspose.Words;
using Aspose.Words.Replacing;

class Program
{
    static void Main()
    {
        // Load the DOCX file.
        Document doc = new Document("Input.docx");

        // Text to find and its replacement.
        string pattern = "_FullName_";
        string replacement = "John Doe";

        // Perform a find‑and‑replace operation on the whole document.
        // This method works on the underlying Run nodes, replacing the text wherever it occurs.
        int replacedCount = doc.Range.Replace(pattern, replacement);

        // Optional: use FindReplaceOptions for more control (e.g., case‑sensitive, whole‑word only).
        // FindReplaceOptions options = new FindReplaceOptions { MatchCase = true, FindWholeWordsOnly = true };
        // int replacedCount = doc.Range.Replace(pattern, replacement, options);

        // Save the updated document.
        doc.Save("Output.docx");
    }
}
