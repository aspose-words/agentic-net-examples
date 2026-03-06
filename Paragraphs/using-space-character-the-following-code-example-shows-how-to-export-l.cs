using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

namespace AsposeWordsListIndentationExample
{
    class Program
    {
        static void Main()
        {
            // Create a new empty document.
            Document doc = new Document();

            // Use DocumentBuilder to add a three‑level numbered list.
            DocumentBuilder builder = new DocumentBuilder(doc);
            builder.ListFormat.ApplyNumberDefault(); // Level 0
            builder.Writeln("Item 1");               // 1. Item 1

            builder.ListFormat.ListIndent();         // Level 1
            builder.Writeln("Item 2");               // a. Item 2

            builder.ListFormat.ListIndent();         // Level 2
            builder.Write("Item 3");                 // i. Item 3

            // Configure TxtSaveOptions to use a space character for list indentation.
            TxtSaveOptions txtSaveOptions = new TxtSaveOptions
            {
                // Use a space as the padding character.
                ListIndentation = { Character = ' ', Count = 3 }
            };

            // Define output paths.
            string artifactsDir = Path.Combine(Environment.CurrentDirectory, "Output");
            Directory.CreateDirectory(artifactsDir);
            string txtPath = Path.Combine(artifactsDir, "ListWithSpaceIndentation.txt");
            string docxPath = Path.Combine(artifactsDir, "ListWithSpaceIndentation.docx");

            // Save the document as plain text using the configured options.
            doc.Save(txtPath, txtSaveOptions);

            // Also save the original document as DOCX for reference.
            doc.Save(docxPath, SaveFormat.Docx);
        }
    }
}
