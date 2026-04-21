using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // Prepare output folder.
        string artifactsDir = Path.Combine(Directory.GetCurrentDirectory(), "Artifacts");
        Directory.CreateDirectory(artifactsDir);

        // Create a sample document with headings.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Heading1;
        builder.Writeln("Chapter One");
        builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Normal;
        builder.Writeln("Content of chapter one.");

        builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Heading2;
        builder.Writeln("Section 1.1");
        builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Normal;
        builder.Writeln("Details of section 1.1.");

        builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Heading1;
        builder.Writeln("Chapter Two");
        builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Normal;
        builder.Writeln("Content of chapter two.");

        // Collect heading texts in the order they appear.
        List<string> headingTexts = doc.GetChildNodes(NodeType.Paragraph, true)
            .Cast<Paragraph>()
            .Where(p => p.ParagraphFormat.IsHeading)
            .Select(p => p.GetText().Trim())
            .ToList();

        // Configure HTML save options to split by heading paragraphs.
        HtmlSaveOptions saveOptions = new HtmlSaveOptions
        {
            DocumentSplitCriteria = DocumentSplitCriteria.HeadingParagraph,
            DocumentSplitHeadingLevel = 2 // split at Heading 1 and Heading 2.
        };

        // Assign the custom callback that renames each part based on the heading text.
        saveOptions.DocumentPartSavingCallback = new HeadingBasedDocumentPartSavingCallback(headingTexts, artifactsDir);

        // Save the document; this will produce several HTML files.
        string mainOutputPath = Path.Combine(artifactsDir, "SplitDocument.html");
        doc.Save(mainOutputPath, saveOptions);

        // Verify that each expected file was created.
        foreach (string heading in headingTexts)
        {
            string safeName = MakeSafeFileName(heading);
            string expectedPath = Path.Combine(artifactsDir, safeName + ".html");
            if (!File.Exists(expectedPath))
                throw new InvalidOperationException($"Expected split file not found: {expectedPath}");
        }
    }

    // Helper to replace invalid filename characters.
    private static string MakeSafeFileName(string name)
    {
        char[] invalid = Path.GetInvalidFileNameChars();
        return string.Concat(name.Select(ch => invalid.Contains(ch) ? '_' : ch));
    }
}

// Callback that assigns filenames based on the original heading text.
public class HeadingBasedDocumentPartSavingCallback : IDocumentPartSavingCallback
{
    private readonly IList<string> _headings;
    private readonly string _outputDir;
    private int _currentIndex = 0;

    public HeadingBasedDocumentPartSavingCallback(IList<string> headings, string outputDir)
    {
        _headings = headings;
        _outputDir = outputDir;
    }

    void IDocumentPartSavingCallback.DocumentPartSaving(DocumentPartSavingArgs args)
    {
        if (_currentIndex < _headings.Count)
        {
            string safeHeading = MakeSafeFileName(_headings[_currentIndex]);
            string newFileName = safeHeading + Path.GetExtension(args.DocumentPartFileName);

            // Set the new filename (without path). Aspose will combine it with the main output directory.
            args.DocumentPartFileName = newFileName;

            // Optionally, provide a stream with a full path.
            string fullPath = Path.Combine(_outputDir, newFileName);
            args.DocumentPartStream = new FileStream(fullPath, FileMode.Create);
        }

        _currentIndex++;
    }

    private static string MakeSafeFileName(string name)
    {
        char[] invalid = Path.GetInvalidFileNameChars();
        return string.Concat(name.Select(ch => invalid.Contains(ch) ? '_' : ch));
    }
}
