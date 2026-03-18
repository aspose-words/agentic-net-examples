using Aspose.Words;
using System;
using System.Linq;

class Program
{
    static void Main()
    {
        // Create a new document.
        Document doc = new Document();

        // Build the document with several paragraphs.
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("Paragraph 1");
        builder.Writeln("Paragraph 2");
        builder.Writeln("Paragraph 3");
        builder.Writeln("Paragraph 4");

        // Access the paragraph collection.
        ParagraphCollection paragraphs = doc.FirstSection.Body.Paragraphs;

        // Ensure the collection has at least two items.
        if (paragraphs.Count >= 2)
        {
            // Calculate the index of the second‑to‑last paragraph.
            int secondToLastIndex = paragraphs.Count - 2;

            // Retrieve the second‑to‑last paragraph using ElementAt and cast to Paragraph.
            Paragraph secondToLast = (Paragraph)paragraphs.ElementAt(secondToLastIndex);

            // Output its text.
            Console.WriteLine("Second‑to‑last paragraph: " + secondToLast.GetText().Trim());
        }

        // Save the document.
        doc.Save("Output.docx");
    }
}
