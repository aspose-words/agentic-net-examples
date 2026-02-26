using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Loading;
using Aspose.Words.Saving;

class Program
{
    static void Main()
    {
        // Sample markdown content
        string markdown = "# Title\n\nThis is a sample markdown document.\n\n[[Placeholder]]";

        // Load the markdown into a Document using MarkdownLoadOptions
        using (MemoryStream inputStream = new MemoryStream(System.Text.Encoding.UTF8.GetBytes(markdown)))
        {
            MarkdownLoadOptions loadOptions = new MarkdownLoadOptions
            {
                PreserveEmptyLines = true // keep empty lines during load
            };

            Document doc = new Document(inputStream, loadOptions);

            // Replace a placeholder with actual text
            doc.Range.Replace("[[Placeholder]]", "Replaced content");

            // Save the document as a DOCX file
            doc.Save("Output.docx", SaveFormat.Docx);

            // Save the document back to Markdown with custom export options
            MarkdownSaveOptions saveOptions = new MarkdownSaveOptions
            {
                ExportAsHtml = MarkdownExportAsHtml.Tables // export tables as raw HTML (example)
            };

            doc.Save("Output.md", saveOptions);
        }
    }
}
