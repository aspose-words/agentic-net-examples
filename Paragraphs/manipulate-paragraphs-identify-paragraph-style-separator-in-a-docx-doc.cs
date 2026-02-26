using System;
using Aspose.Words;

namespace ParagraphStyleSeparatorDemo
{
    class Program
    {
        static void Main()
        {
            // Load an existing DOCX document.
            // The Document constructor is the approved way to create/load a document.
            Document doc = new Document("Input.docx");

            // Iterate through all paragraphs in the main body of the first section.
            // ParagraphCollection implements IEnumerable, so we can use a foreach loop.
            int paragraphIndex = 0;
            foreach (Paragraph para in doc.FirstSection.Body.Paragraphs)
            {
                // The BreakIsStyleSeparator property tells us whether this paragraph ends
                // with a style separator (the special character that allows mixed styles in one line).
                if (para.BreakIsStyleSeparator)
                {
                    Console.WriteLine($"Paragraph at index {paragraphIndex} contains a style separator.");
                }

                paragraphIndex++;
            }

            // Optionally, save the document after processing (no changes made here).
            // The Save method with a file name is the approved way to persist a document.
            doc.Save("Output.docx");
        }
    }
}
