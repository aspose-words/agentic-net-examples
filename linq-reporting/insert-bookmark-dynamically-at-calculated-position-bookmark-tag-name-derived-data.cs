using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Words;
using Aspose.Words.Tables;

namespace BookmarkInsertionDemo
{
    // Simple data model whose fields will be used to build bookmark names.
    public class DataItem
    {
        public int Id { get; set; }
        public string Name { get; set; }
    }

    public static class BookmarkHelper
    {
        /// <summary>
        /// Loads a Word document (or creates a new one if the file does not exist),
        /// inserts a bookmark for each data item at a calculated position,
        /// and saves the result.
        /// </summary>
        /// <param name="inputPath">Path to the source .docx file.</param>
        /// <param name="outputPath">Path where the modified document will be saved.</param>
        /// <param name="items">Collection of data items used to generate bookmark names.</param>
        public static void InsertBookmarks(string inputPath, string outputPath, List<DataItem> items)
        {
            // Load the existing document or create a new one if the file is missing.
            Document doc = File.Exists(inputPath) ? new Document(inputPath) : new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Iterate over the data items and create a bookmark for each.
            foreach (DataItem item in items)
            {
                // Derive a unique bookmark name from the data fields.
                // Example: "Bookmark_12_JohnDoe"
                string bookmarkName = $"Bookmark_{item.Id}_{item.Name.Replace(' ', '_')}";

                // Calculate the insertion point.
                // For demonstration we move to the end of the document,
                // insert a paragraph break, then place the bookmark.
                builder.MoveToDocumentEnd();
                builder.Writeln(); // Ensure we start on a new paragraph.

                // Start the bookmark, write some content, and close the bookmark.
                builder.StartBookmark(bookmarkName);
                builder.Write($"Data Item: Id={item.Id}, Name={item.Name}");
                builder.EndBookmark(bookmarkName);
            }

            // Example of moving to a specific bookmark that was just created
            // and inserting additional text after it.
            if (items.Count > 0)
            {
                string firstBookmark = $"Bookmark_{items[0].Id}_{items[0].Name.Replace(' ', '_')}";
                // Move to the start of the first bookmark (precision: start, after the start node).
                if (builder.MoveToBookmark(firstBookmark, true, true))
                {
                    builder.Writeln(); // Add a new line after the bookmark start.
                    builder.Write("[Additional info after first bookmark]");
                }
            }

            // Save the modified document.
            doc.Save(outputPath);
        }

        // Example usage.
        public static void Main()
        {
            // Prepare sample data.
            var data = new List<DataItem>
            {
                new DataItem { Id = 1, Name = "Alice Smith" },
                new DataItem { Id = 2, Name = "Bob Johnson" },
                new DataItem { Id = 3, Name = "Charlie Brown" }
            };

            // Use temporary paths that are guaranteed to exist.
            string tempFolder = Path.GetTempPath();
            string inputFile = Path.Combine(tempFolder, "Template.docx");
            string outputFile = Path.Combine(tempFolder, "Result.docx");

            // Ensure an input file exists; if not, create an empty document.
            if (!File.Exists(inputFile))
            {
                new Document().Save(inputFile);
            }

            // Perform the bookmark insertion.
            InsertBookmarks(inputFile, outputFile, data);

            Console.WriteLine($"Bookmarks inserted and document saved to: {outputFile}");
        }
    }
}
