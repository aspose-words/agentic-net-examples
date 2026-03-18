using System;
using System.Collections.Generic;
using Aspose.Words;

namespace BookmarkInsertionExample
{
    class Program
    {
        static void Main()
        {
            // Create a new empty document.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Sample list of expressions – only non‑empty strings will be turned into bookmarks.
            List<string> expressions = new List<string>
            {
                "FirstBookmark",
                "",                     // Empty – should be ignored.
                "SecondBookmark",
                "   ",                  // Whitespace – should be ignored.
                "ThirdBookmark"
            };

            // Iterate over each expression and insert a bookmark with a placeholder tag.
            foreach (string expr in expressions)
            {
                // Skip null, empty or whitespace‑only expressions.
                if (string.IsNullOrWhiteSpace(expr))
                    continue;

                // Start the bookmark using the expression as the bookmark name.
                builder.StartBookmark(expr);

                // Insert the literal placeholder tag that references the bookmark.
                // The format required is: <<bookmark [Expression]>>
                builder.Write($"<<bookmark {expr}>>");

                // End the bookmark with the same name.
                builder.EndBookmark(expr);

                // Add a line break after each placeholder for readability.
                builder.Writeln();
            }

            // Save the document to verify the result.
            doc.Save("BookmarksWithTags.docx");
        }
    }
}
