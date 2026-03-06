using System;
using Aspose.Words;
using Aspose.Words.Loading;
using Aspose.Words.Saving;

class Program
{
    static void Main()
    {
        // Paths to the MHTML template, source documents and the output file.
        string templatePath = "Template.mht";
        string[] sourceDocs = { "Part1.docx", "Part2.docx", "Part3.docx" };
        string outputPath = "Result.docx";

        // Load the MHTML template. Use an object initializer to set LoadFormat to Mhtml.
        LoadOptions loadOptions = new LoadOptions { LoadFormat = LoadFormat.Mhtml };
        Document template = new Document(templatePath, loadOptions);

        // DocumentBuilder will be used to navigate the template and insert other documents.
        DocumentBuilder builder = new DocumentBuilder(template);

        // The template should contain bookmarks named InsertHere1, InsertHere2, … where the
        // corresponding source documents will be placed.
        for (int i = 0; i < sourceDocs.Length; i++)
        {
            string bookmarkName = $"InsertHere{i + 1}";

            // Verify that the bookmark exists before trying to move to it.
            if (template.Range.Bookmarks[bookmarkName] != null)
            {
                // Position the cursor at the bookmark.
                builder.MoveToBookmark(bookmarkName);

                // Load the source document that will be inserted.
                Document srcDoc = new Document(sourceDocs[i]);

                // Insert the source document at the current cursor position.
                // KeepSourceFormatting preserves the original formatting of the inserted content.
                builder.InsertDocument(srcDoc, ImportFormatMode.KeepSourceFormatting);
            }
            else
            {
                Console.WriteLine($"Bookmark '{bookmarkName}' not found in the template.");
            }
        }

        // Save the combined document. The SaveFormat is inferred from the file extension,
        // but we explicitly specify Docx for clarity.
        template.Save(outputPath, SaveFormat.Docx);
    }
}
