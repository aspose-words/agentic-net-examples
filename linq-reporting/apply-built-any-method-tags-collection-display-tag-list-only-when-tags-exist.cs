using System;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Markup;

class Program
{
    static void Main()
    {
        // Load an existing document (replace with your actual path)
        Document doc = new Document("Input.docx");

        // Get the collection of structured document tags in the document.
        StructuredDocumentTagCollection tags = doc.Range.StructuredDocumentTags;

        // Display the list of tags only if the collection contains any items.
        if (tags.Any())
        {
            Console.WriteLine("Structured Document Tags found:");
            foreach (IStructuredDocumentTag tag in tags)
            {
                // Prefer the Title if set; otherwise fall back to the Tag string.
                string displayName = !string.IsNullOrEmpty(tag.Title) ? tag.Title : tag.Tag;
                Console.WriteLine($"- {displayName}");
            }
        }
        else
        {
            Console.WriteLine("No structured document tags present in the document.");
        }

        // Save the document (replace with your desired output path)
        doc.Save("Output.docx");
    }
}
