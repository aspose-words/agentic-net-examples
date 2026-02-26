using System;
using Aspose.Words;
using Aspose.Words.Vba;
using Aspose.Words.Tables;
using Aspose.Words.Notes; // Added for Footnote class

namespace AsposeWordsVbaDemo
{
    class Program
    {
        static void Main()
        {
            // Load an existing Word document (DOCM to ensure it may contain macros).
            // This uses the Document(string) constructor – the approved lifecycle rule for loading.
            Document doc = new Document("InputDocument.docm");

            // -----------------------------------------------------------------
            // 1. Check whether the document contains VBA macros.
            // -----------------------------------------------------------------
            Console.WriteLine($"Document has macros: {doc.HasMacros}");

            if (doc.HasMacros && doc.VbaProject != null)
            {
                // List all VBA modules and their source code.
                VbaProject vbaProject = doc.VbaProject;
                Console.WriteLine($"VBA Project Name: {vbaProject.Name}");
                Console.WriteLine($"Number of modules: {vbaProject.Modules.Count}");

                foreach (VbaModule module in vbaProject.Modules)
                {
                    Console.WriteLine($"--- Module: {module.Name} ---");
                    Console.WriteLine(module.SourceCode);
                }
            }

            // -----------------------------------------------------------------
            // 2. Extract text that lies inside a bookmark named "MyBookmark".
            // -----------------------------------------------------------------
            string bookmarkText = string.Empty;
            if (doc.Range.Bookmarks["MyBookmark"] != null)
            {
                // Bookmark.Text returns the text contained within the bookmark.
                bookmarkText = doc.Range.Bookmarks["MyBookmark"].Text;
                Console.WriteLine($"Text inside bookmark 'MyBookmark': {bookmarkText}");
            }
            else
            {
                Console.WriteLine("Bookmark 'MyBookmark' not found.");
            }

            // -----------------------------------------------------------------
            // 3. Access header and footer text of the first section.
            // -----------------------------------------------------------------
            HeaderFooter header = doc.FirstSection.HeadersFooters[HeaderFooterType.HeaderPrimary];
            HeaderFooter footer = doc.FirstSection.HeadersFooters[HeaderFooterType.FooterPrimary];

            string headerText = header != null ? header.GetText().Trim() : string.Empty;
            string footerText = footer != null ? footer.GetText().Trim() : string.Empty;

            Console.WriteLine($"Header text: {headerText}");
            Console.WriteLine($"Footer text: {footerText}");

            // -----------------------------------------------------------------
            // 4. Retrieve all footnotes in the document.
            // -----------------------------------------------------------------
            NodeCollection footnoteNodes = doc.GetChildNodes(NodeType.Footnote, true);
            Console.WriteLine($"Total footnotes: {footnoteNodes.Count}");

            for (int i = 0; i < footnoteNodes.Count; i++)
            {
                Footnote footnote = (Footnote)footnoteNodes[i];
                Console.WriteLine($"Footnote {i + 1}: {footnote.GetText().Trim()}");
            }

            // -----------------------------------------------------------------
            // 5. Create a new document that summarizes the extracted information.
            // -----------------------------------------------------------------
            // This uses the Document() constructor – the approved lifecycle rule for creation.
            Document summaryDoc = new Document();
            DocumentBuilder builder = new DocumentBuilder(summaryDoc);

            builder.Writeln("=== Document Summary ===");
            builder.Writeln($"Has Macros: {doc.HasMacros}");
            builder.Writeln();

            if (doc.HasMacros && doc.VbaProject != null)
            {
                builder.Writeln("=== VBA Project ===");
                builder.Writeln($"Project Name: {doc.VbaProject.Name}");
                builder.Writeln($"Modules Count: {doc.VbaProject.Modules.Count}");
                foreach (VbaModule module in doc.VbaProject.Modules)
                {
                    builder.Writeln($"--- Module: {module.Name} ---");
                    builder.Writeln(module.SourceCode);
                }
                builder.Writeln();
            }

            builder.Writeln("=== Bookmark Content ===");
            builder.Writeln(string.IsNullOrEmpty(bookmarkText)
                ? "Bookmark 'MyBookmark' not found."
                : bookmarkText);
            builder.Writeln();

            builder.Writeln("=== Header & Footer ===");
            builder.Writeln($"Header: {headerText}");
            builder.Writeln($"Footer: {footerText}");
            builder.Writeln();

            builder.Writeln("=== Footnotes ===");
            for (int i = 0; i < footnoteNodes.Count; i++)
            {
                Footnote footnote = (Footnote)footnoteNodes[i];
                builder.Writeln($"Footnote {i + 1}: {footnote.GetText().Trim()}");
            }

            // Save the summary document.
            // This uses the Document.Save(string) method – the approved lifecycle rule for saving.
            summaryDoc.Save("DocumentSummary.docx");

            Console.WriteLine("Summary document saved as 'DocumentSummary.docx'.");
        }
    }
}
