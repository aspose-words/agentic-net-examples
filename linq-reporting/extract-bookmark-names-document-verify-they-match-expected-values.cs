using System;
using System.Collections.Generic;
using Aspose.Words;

class BookmarkVerification
{
    static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert three bookmarks with known names.
        for (int i = 1; i <= 3; i++)
        {
            string name = $"MyBookmark_{i}";
            builder.StartBookmark(name);
            builder.Write($"Text inside {name}.");
            builder.EndBookmark(name);
            builder.Writeln(); // separate paragraphs
        }

        // Expected bookmark names.
        string[] expected = { "MyBookmark_1", "MyBookmark_2", "MyBookmark_3" };

        // Extract actual bookmark names from the document.
        BookmarkCollection bookmarks = doc.Range.Bookmarks;
        List<string> actual = new List<string>();
        foreach (Bookmark bm in bookmarks)
        {
            actual.Add(bm.Name);
        }

        // Verify that the extracted names match the expected names.
        bool match = expected.Length == actual.Count;
        if (match)
        {
            for (int i = 0; i < expected.Length; i++)
            {
                if (!string.Equals(expected[i], actual[i], StringComparison.OrdinalIgnoreCase))
                {
                    match = false;
                    break;
                }
            }
        }

        Console.WriteLine(match
            ? "All bookmark names match the expected values."
            : "Bookmark names do NOT match the expected values.");

        // Optionally, display the extracted names.
        Console.WriteLine("Extracted bookmark names:");
        foreach (string name in actual)
        {
            Console.WriteLine(name);
        }
    }
}
