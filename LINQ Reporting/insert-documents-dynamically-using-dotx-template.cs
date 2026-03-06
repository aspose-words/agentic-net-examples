using System;
using System.Collections.Generic;
using Aspose.Words;
using Aspose.Words.Saving;

namespace AsposeWordsDynamicInsert
{
    class Program
    {
        static void Main()
        {
            // Path to the DOTX template that contains placeholders or bookmarks.
            string templatePath = @"C:\Docs\Template.dotx";

            // Load the DOTX template. The Document constructor creates the document object.
            Document templateDoc = new Document(templatePath);

            // Create a DocumentBuilder associated with the template.
            DocumentBuilder builder = new DocumentBuilder(templateDoc);

            // Move the cursor to the end of the document (or to a specific bookmark if required).
            builder.MoveToDocumentEnd();

            // List of source documents (DOCX, DOTX, etc.) that will be inserted dynamically.
            List<string> sourceDocs = new List<string>
            {
                @"C:\Docs\Section1.docx",
                @"C:\Docs\Section2.docx",
                @"C:\Docs\Section3.docx"
            };

            // Insert each source document at the current cursor position.
            foreach (string srcPath in sourceDocs)
            {
                // Load the source document.
                Document srcDoc = new Document(srcPath);

                // Insert a page break before each inserted document for visual separation.
                builder.InsertBreak(BreakType.PageBreak);

                // Insert the source document. KeepSourceFormatting preserves the original styles.
                builder.InsertDocument(srcDoc, ImportFormatMode.KeepSourceFormatting);
            }

            // Save the final assembled document.
            string outputPath = @"C:\Docs\Result.docx";
            templateDoc.Save(outputPath, SaveFormat.Docx);
        }
    }
}
