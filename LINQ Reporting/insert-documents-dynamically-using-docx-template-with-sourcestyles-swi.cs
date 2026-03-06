using System;
using System.Collections.Generic;
using Aspose.Words;
using Aspose.Words.Reporting;

namespace AsposeWordsDynamicInsert
{
    class Program
    {
        static void Main()
        {
            // Path to the template document that contains bookmarks where other docs will be inserted.
            string templatePath = @"C:\Docs\Template.docx";

            // Load the template.
            Document templateDoc = new Document(templatePath);
            DocumentBuilder builder = new DocumentBuilder(templateDoc);

            // List of source documents to insert.
            List<string> sourcePaths = new List<string>
            {
                @"C:\Docs\Part1.docx",
                @"C:\Docs\Part2.docx",
                @"C:\Docs\Part3.docx"
            };

            // Flag that determines whether to keep source formatting (sourceStyles switch).
            bool useSourceStyles = true;

            // Prepare import options. No need to set SmartStyleBehavior when we already specify ImportFormatMode.
            ImportFormatOptions importOptions = new ImportFormatOptions();

            // Choose the appropriate ImportFormatMode based on the flag.
            ImportFormatMode importMode = useSourceStyles
                ? ImportFormatMode.KeepSourceFormatting
                : ImportFormatMode.UseDestinationStyles;

            // Insert each source document at a corresponding bookmark.
            // Bookmarks in the template are expected to be named "InsertHere1", "InsertHere2", etc.
            for (int i = 0; i < sourcePaths.Count; i++)
            {
                string bookmarkName = $"InsertHere{i + 1}";

                // Verify that the bookmark exists. BookmarkCollection does not have an Exists method;
                // we check the indexer for a null result instead.
                if (templateDoc.Range.Bookmarks[bookmarkName] == null)
                    continue; // Skip if the bookmark does not exist.

                // Load the source document.
                Document srcDoc = new Document(sourcePaths[i]);

                // Move the builder cursor to the bookmark.
                builder.MoveToBookmark(bookmarkName);

                // Insert the source document using the selected import mode and options.
                builder.InsertDocument(srcDoc, importMode, importOptions);
            }

            // Save the resulting document.
            string outputPath = @"C:\Docs\Result.docx";
            templateDoc.Save(outputPath);
        }
    }
}
