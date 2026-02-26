using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Markup;
using Aspose.Words.Saving;
using Aspose.Words.Loading;

class Program
{
    static void Main()
    {
        // -----------------------------------------------------------------
        // 1. Create a new blank document and insert a content control (SDT).
        // -----------------------------------------------------------------
        Document doc = new Document();                     // Create a blank document.
        DocumentBuilder builder = new DocumentBuilder(doc); // Helper to add content.

        // Create an inline plain‑text content control.
        StructuredDocumentTag sdt = new StructuredDocumentTag(
            doc,
            SdtType.PlainText,          // Plain text content control.
            MarkupLevel.Inline)         // Occurs at the inline level.
        {
            Title = "SampleControl",    // Friendly name.
            Tag = "SampleTag",          // Tag identifier.
            LockContentControl = true, // Prevent deletion by the user.
            LockContents = false       // Allow editing of the text inside.
        };

        // Insert the content control into the document.
        builder.InsertNode(sdt);

        // Move the cursor inside the newly inserted content control and add some text.
        builder.MoveTo(sdt);
        builder.Write("Initial content inside the control.");

        // Add a paragraph after the control for clarity.
        builder.Writeln();
        builder.Writeln("Following paragraph after the content control.");

        // ---------------------------------------------------------------
        // 2. Save the document as Markdown with custom save options.
        // ---------------------------------------------------------------
        MarkdownSaveOptions saveOptions = new MarkdownSaveOptions
        {
            // Export any tables as raw HTML (not used here but shown as an example).
            ExportAsHtml = MarkdownExportAsHtml.Tables,
            // Ensure the generator name is included in the output.
            ExportGeneratorName = true,
            // Use UTF‑8 encoding.
            Encoding = System.Text.Encoding.UTF8
        };

        string markdownPath = Path.Combine(Environment.CurrentDirectory, "SampleDocument.md");
        doc.Save(markdownPath, saveOptions);

        // ---------------------------------------------------------------
        // 3. Load the previously saved Markdown document preserving empty lines.
        // ---------------------------------------------------------------
        MarkdownLoadOptions loadOptions = new MarkdownLoadOptions
        {
            PreserveEmptyLines = true // Keep empty lines from the original Markdown.
        };

        Document loadedDoc = new Document(markdownPath, loadOptions);

        // ---------------------------------------------------------------
        // 4. Iterate over all content controls in the loaded document and
        //    modify their properties (e.g., make them read‑only).
        // ---------------------------------------------------------------
        NodeCollection sdtNodes = loadedDoc.GetChildNodes(NodeType.StructuredDocumentTag, true);
        foreach (StructuredDocumentTag tag in sdtNodes)
        {
            // Make the content control read‑only.
            tag.LockContents = true;
        }

        // ---------------------------------------------------------------
        // 5. Save the modified document back to Markdown.
        // ---------------------------------------------------------------
        string modifiedMarkdownPath = Path.Combine(Environment.CurrentDirectory, "ModifiedSampleDocument.md");
        loadedDoc.Save(modifiedMarkdownPath, saveOptions);

        // Inform the user that the process has completed.
        Console.WriteLine("Original Markdown saved to: " + markdownPath);
        Console.WriteLine("Modified Markdown saved to: " + modifiedMarkdownPath);
    }
}
