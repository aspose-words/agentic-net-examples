using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;
using Aspose.Words.Fields;

namespace RemoveTocAndSaveAsMarkdown
{
    class Program
    {
        static void Main()
        {
            // Path to the source DOC document.
            string inputPath = @"C:\Docs\SourceDocument.docx";

            // Path where the resulting Markdown file will be saved.
            string outputPath = @"C:\Docs\ResultDocument.md";

            // Load the existing document from file.
            Document doc = new Document(inputPath);

            // Remove all Table of Contents (TOC) fields from the document.
            // Iterate over a copy of the fields collection because removing a field modifies the collection.
            foreach (Field field in doc.Range.Fields.ToArray())
            {
                if (field.Type == FieldType.FieldTOC)
                {
                    // The Remove method returns the node that follows the removed field,
                    // but we do not need to use the return value here.
                    field.Remove();
                }
            }

            // Configure Markdown save options if needed (default options are sufficient for this task).
            MarkdownSaveOptions saveOptions = new MarkdownSaveOptions
            {
                // Ensure the format is explicitly set to Markdown.
                SaveFormat = SaveFormat.Markdown
            };

            // Save the modified document as a Markdown file.
            doc.Save(outputPath, saveOptions);
        }
    }
}
