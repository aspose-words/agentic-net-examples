using System;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Vba;
using Aspose.Words.Notes;

class Program
{
    static void Main()
    {
        // Load an existing DOCX file.
        Document doc = new Document("Input.docx");

        // ----- Extract main document text (the whole story) -----
        string mainText = doc.Range.Text;
        Console.WriteLine("=== Main Document Text ===");
        Console.WriteLine(mainText);

        // ----- Extract text inside a specific bookmark (range) -----
        // Assumes a bookmark named "MyRange" exists in the document.
        if (doc.Range.Bookmarks["MyRange"] != null)
        {
            string bookmarkText = doc.Range.Bookmarks["MyRange"].Text;
            Console.WriteLine("\n=== Bookmark 'MyRange' Text ===");
            Console.WriteLine(bookmarkText);
        }

        // ----- Access headers and footers for each section -----
        foreach (Section section in doc.Sections)
        {
            // Primary header (odd pages or default header).
            HeaderFooter primaryHeader = section.HeadersFooters[HeaderFooterType.HeaderPrimary];
            if (primaryHeader != null)
            {
                Console.WriteLine("\n=== Primary Header ===");
                Console.WriteLine(primaryHeader.GetText());
            }

            // Primary footer (odd pages or default footer).
            HeaderFooter primaryFooter = section.HeadersFooters[HeaderFooterType.FooterPrimary];
            if (primaryFooter != null)
            {
                Console.WriteLine("\n=== Primary Footer ===");
                Console.WriteLine(primaryFooter.GetText());
            }
        }

        // ----- Access footnotes -----
        // Retrieve all footnote nodes in the document.
        NodeCollection footnoteNodes = doc.GetChildNodes(NodeType.Footnote, true);
        foreach (Footnote footnote in footnoteNodes)
        {
            Console.WriteLine("\n=== Footnote ===");
            Console.WriteLine(footnote.GetText());
        }

        // ----- If the document contains VBA macros, list them -----
        if (doc.HasMacros)
        {
            VbaProject vbaProject = doc.VbaProject;
            Console.WriteLine($"\n=== VBA Project: {vbaProject.Name} (Modules: {vbaProject.Modules.Count()}) ===");
            foreach (VbaModule module in vbaProject.Modules)
            {
                Console.WriteLine($"\n--- Module: {module.Name} ---");
                Console.WriteLine(module.SourceCode);
            }
        }
    }
}
