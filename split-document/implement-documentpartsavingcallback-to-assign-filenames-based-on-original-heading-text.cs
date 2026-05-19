using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // Prepare output folder.
        string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "Output");
        Directory.CreateDirectory(outputDir);

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
        List<string> headings = new List<string>();
        foreach (Paragraph para in doc.GetChildNodes(NodeType.Paragraph, true))
        {
            if (para.ParagraphFormat.IsHeading)
                headings.Add(para.GetText().Trim());
        }

        // Configure HTML save options to split by heading paragraphs.
        HtmlSaveOptions saveOptions = new HtmlSaveOptions
        {
            DocumentSplitCriteria = DocumentSplitCriteria.HeadingParagraph,
            DocumentSplitHeadingLevel = 2 // split at Heading 1 and Heading 2.
        };

        // Assign the custom callback that names each part after its heading text.
        saveOptions.DocumentPartSavingCallback = new HeadingBasedDocumentPartSavingCallback(headings, outputDir);

        // Save the document; it will be split into multiple HTML files.
        string mainFilePath = Path.Combine(outputDir, "Combined.html");
        doc.Save(mainFilePath, saveOptions);

        // Verify that each expected part file exists.
        foreach (string heading in headings)
        {
            string safeName = MakeFileNameSafe(heading) + ".html";
            string partPath = Path.Combine(outputDir, safeName);
            if (!File.Exists(partPath))
                throw new FileNotFoundException($"Expected split part not found: {partPath}");
        }

        // Indicate successful completion.
        Console.WriteLine("Document split completed. Files are located in: " + outputDir);
    }

    // Helper to replace invalid filename characters.
    private static string MakeFileNameSafe(string name)
    {
        foreach (char c in Path.GetInvalidFileNameChars())
            name = name.Replace(c, '_');
        return name;
    }
}

// Callback that assigns filenames based on the original heading text.
public class HeadingBasedDocumentPartSavingCallback : IDocumentPartSavingCallback
{
    private readonly List<string> _headings;
    private readonly string _outputDir;
    private int _currentIndex = 0;

    public HeadingBasedDocumentPartSavingCallback(List<string> headings, string outputDir)
    {
        _headings = headings;
        _outputDir = outputDir;
    }

    void IDocumentPartSavingCallback.DocumentPartSaving(DocumentPartSavingArgs args)
    {
        // Determine the heading text for this part.
        string heading = _currentIndex < _headings.Count ? _headings[_currentIndex] : $"Part{_currentIndex + 1}";
        string safeHeading = MakeFileNameSafe(heading);
        string extension = Path.GetExtension(args.DocumentPartFileName);
        string fileName = $"{safeHeading}{extension}";

        // Set the new filename and stream for the part.
        args.DocumentPartFileName = fileName;
        args.DocumentPartStream = new FileStream(Path.Combine(_outputDir, fileName), FileMode.Create);
        _currentIndex++;
    }

    private static string MakeFileNameSafe(string name)
    {
        foreach (char c in Path.GetInvalidFileNameChars())
            name = name.Replace(c, '_');
        return name;
    }
}
