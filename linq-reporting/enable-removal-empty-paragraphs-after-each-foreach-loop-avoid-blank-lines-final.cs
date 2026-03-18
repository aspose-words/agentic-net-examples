using System;
using System.Collections.Generic;
using Aspose.Words;
using Aspose.Words.MailMerging;

class RemoveEmptyParagraphsDemo
{
    static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Sample data that may contain empty strings.
        List<string> lines = new List<string> { "First line", "", "Third line", "   ", "Fifth line" };

        // Iterate over the data and write each item as a separate paragraph.
        foreach (string line in lines)
        {
            // Write the current line (may be empty or whitespace).
            builder.Writeln(line);

            // Configure mail‑merge cleanup to remove empty paragraphs.
            // This option removes any paragraph that contains no visible content.
            doc.MailMerge.CleanupOptions = MailMergeCleanupOptions.RemoveEmptyParagraphs;

            // Execute a dummy mail merge (no fields) to trigger the cleanup logic.
            // The empty field arrays mean no actual merge occurs; only cleanup is applied.
            doc.MailMerge.Execute(new string[0], new object[0]);
        }

        // Save the resulting document. Empty paragraphs introduced by the loop are removed.
        doc.Save("Result.docx");
    }
}
