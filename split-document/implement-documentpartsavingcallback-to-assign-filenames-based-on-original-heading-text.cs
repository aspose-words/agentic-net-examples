using System;
using System.IO;
using System.Collections.Generic;
using Aspose.Words;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // Prepare output directory.
        string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "Output");
        if (Directory.Exists(outputDir))
            Directory.Delete(outputDir, true);
        Directory.CreateDirectory(outputDir);

        // Create a sample document with headings.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Heading1;
        builder.Writeln("Chapter One");
        builder.Writeln("Content of chapter one.");

        builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Heading2;
        builder.Writeln("Section 1.1");
        builder.Writeln("Details of section 1.1.");

        builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Heading1;
        builder.Writeln("Chapter Two");
        builder.Writeln("Content of chapter two.");

        // Collect heading texts in the order they appear.
        List<string> headingTexts = new List<string>();
        foreach (Paragraph para in doc.GetChildNodes(NodeType.Paragraph, true))
        {
            if (para.ParagraphFormat.IsHeading)
                headingTexts.Add(para.GetText().Trim());
        }

        // Configure HTML save options to split by heading paragraphs.
        HtmlSaveOptions saveOptions = new HtmlSaveOptions
        {
            DocumentSplitCriteria = DocumentSplitCriteria.HeadingParagraph,
            DocumentSplitHeadingLevel = 9 // split at all heading levels
        };

        // Assign custom callback to name each part based on heading text.
        saveOptions.DocumentPartSavingCallback = new HeadingBasedPartNaming(outputDir, headingTexts);

        // Save the document; parts will be written by the callback.
        string mainFileName = Path.Combine(outputDir, "Document.html");
        doc.Save(mainFileName, saveOptions);

        // Verify that each expected part file exists.
        foreach (string heading in headingTexts)
        {
            string safeName = MakeSafeFileName(heading) + ".html";
            string partPath = Path.Combine(outputDir, safeName);
            if (!File.Exists(partPath))
                throw new InvalidOperationException($"Expected part file not found: {partPath}");
        }
    }

    // Helper to replace invalid filename characters.
    private static string MakeSafeFileName(string name)
    {
        foreach (char c in Path.GetInvalidFileNameChars())
            name = name.Replace(c, '_');
        return name;
    }

    // Callback implementation that names each document part using the corresponding heading text.
    private class HeadingBasedPartNaming : IDocumentPartSavingCallback
    {
        private readonly string _outputDir;
        private readonly IList<string> _headings;
        private int _index = -1;

        public HeadingBasedPartNaming(string outputDir, IList<string> headings)
        {
            _outputDir = outputDir;
            _headings = headings;
        }

        void IDocumentPartSavingCallback.DocumentPartSaving(DocumentPartSavingArgs args)
        {
            // Increment part counter.
            _index++;

            // Guard against mismatched counts.
            if (_index >= _headings.Count)
                throw new InvalidOperationException("More document parts than headings.");

            // Build a safe filename from the heading text.
            string safeName = MakeSafeFileName(_headings[_index]) + Path.GetExtension(args.DocumentPartFileName);
            string fullPath = Path.Combine(_outputDir, safeName);

            // Set the filename and stream for the part.
            args.DocumentPartFileName = safeName;
            args.DocumentPartStream = new FileStream(fullPath, FileMode.Create);
            args.KeepDocumentPartStreamOpen = false;
        }

        // Reuse the same helper for safety.
        private static string MakeSafeFileName(string name)
        {
            foreach (char c in Path.GetInvalidFileNameChars())
                name = name.Replace(c, '_');
            return name;
        }
    }
}
