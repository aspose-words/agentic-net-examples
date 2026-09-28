using System;
using System.IO;
using System.Collections.Concurrent;
using System.Text;
using System.Threading.Tasks;
using Aspose.Words;
using Aspose.Words.Tables;
using Newtonsoft.Json;

public class Program
{
    public static void Main()
    {
        // Prepare directories
        string baseDir = Path.Combine(Directory.GetCurrentDirectory(), "ExtractionDemo");
        string inputDir = Path.Combine(baseDir, "Input");
        string outputDir = Path.Combine(baseDir, "Output");
        Directory.CreateDirectory(inputDir);
        Directory.CreateDirectory(outputDir);

        // Create sample documents
        int documentCount = 3;
        for (int i = 1; i <= documentCount; i++)
        {
            CreateSampleDocument(Path.Combine(inputDir, $"doc{i}.docx"), i);
        }

        // Collection for extraction results
        var results = new ConcurrentBag<ExtractionResult>();

        // Parallel processing of documents
        string[] files = Directory.GetFiles(inputDir, "*.docx");
        Parallel.ForEach(files, filePath =>
        {
            Document doc = new Document(filePath);
            string extractedText = ExtractTextBetweenBookmarks(doc, "Start", "End");

            string fileName = Path.GetFileNameWithoutExtension(filePath);
            string outPath = Path.Combine(outputDir, $"{fileName}_extracted.txt");
            File.WriteAllText(outPath, extractedText ?? string.Empty);

            if (!File.Exists(outPath))
                throw new InvalidOperationException($"Failed to create output file {outPath}");

            results.Add(new ExtractionResult
            {
                DocumentName = fileName,
                ExtractedText = extractedText ?? string.Empty,
                OutputPath = outPath
            });
        });

        // Write summary JSON
        string summaryPath = Path.Combine(outputDir, "summary.json");
        string json = JsonConvert.SerializeObject(results, Formatting.Indented);
        File.WriteAllText(summaryPath, json);
        if (!File.Exists(summaryPath))
            throw new InvalidOperationException("Summary JSON file was not created.");
    }

    private static void CreateSampleDocument(string path, int index)
    {
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        builder.Writeln($"Document {index} - Intro paragraph.");

        builder.StartBookmark("Start");
        builder.Writeln($"Document {index} - Content inside Start bookmark.");
        builder.EndBookmark("Start");

        builder.StartBookmark("End");
        builder.Writeln($"Document {index} - Content inside End bookmark.");
        builder.EndBookmark("End");

        builder.Writeln($"Document {index} - Closing paragraph.");

        doc.Save(path);
        if (!File.Exists(path))
            throw new InvalidOperationException($"Failed to create sample document {path}");
    }

    private static string ExtractTextBetweenBookmarks(Document doc, string startBookmarkName, string endBookmarkName)
    {
        Bookmark start = doc.Range.Bookmarks[startBookmarkName];
        Bookmark end = doc.Range.Bookmarks[endBookmarkName];
        if (start == null || end == null)
            throw new InvalidOperationException("Required bookmarks not found.");

        // Collect text from nodes that appear after the start bookmark and before the end bookmark.
        Node current = start.BookmarkEnd.NextSibling;
        var sb = new StringBuilder();

        while (current != null && current != end.BookmarkStart)
        {
            if (current.NodeType == NodeType.Paragraph)
                sb.Append(((Paragraph)current).GetText());
            else if (current.NodeType == NodeType.Table)
                sb.Append(((Table)current).GetText());

            current = current.NextSibling;
        }

        return sb.ToString();
    }

    private class ExtractionResult
    {
        public string DocumentName { get; set; }
        public string ExtractedText { get; set; }
        public string OutputPath { get; set; }
    }
}
