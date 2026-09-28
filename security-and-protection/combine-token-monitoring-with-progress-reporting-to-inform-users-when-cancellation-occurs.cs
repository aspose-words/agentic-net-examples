using System;
using System.IO;
using System.Threading;
using System.Threading.Tasks;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Create a sample document with many paragraphs.
        var doc = new Document();
        var builder = new DocumentBuilder(doc);
        for (int i = 0; i < 500; i++)
        {
            builder.Writeln($"Paragraph {i + 1}");
        }

        string sourcePath = "source.docx";
        doc.Save(sourcePath);

        // Load the document for processing.
        var loadedDoc = new Document(sourcePath);

        // Set up cancellation to occur after a short delay.
        var cts = new CancellationTokenSource();
        Task.Delay(100).ContinueWith(_ => cts.Cancel());

        // Progress reporter that writes percentage to the console.
        var progress = new Progress<double>(p => Console.WriteLine($"Progress: {p:P0}"));

        try
        {
            // Process the document while monitoring cancellation and reporting progress.
            ProcessDocument(loadedDoc, cts.Token, progress);

            // Save the processed document.
            string outputPath = "processed.docx";
            loadedDoc.Save(outputPath);

            // Validate that the output file was created.
            if (!File.Exists(outputPath))
                throw new InvalidOperationException($"Failed to create output file: {outputPath}");

            Console.WriteLine($"Document saved to {outputPath}");
        }
        catch (OperationCanceledException)
        {
            Console.WriteLine("Processing was cancelled by the token.");
        }
    }

    private static void ProcessDocument(Document doc, CancellationToken token, IProgress<double> progress)
    {
        // Retrieve all paragraphs in the document.
        var paragraphs = doc.GetChildNodes(NodeType.Paragraph, true);
        int total = paragraphs.Count;

        for (int i = 0; i < total; i++)
        {
            // Throw if cancellation has been requested.
            token.ThrowIfCancellationRequested();

            // Simulate work: append a marker to each paragraph.
            var paragraph = (Paragraph)paragraphs[i];
            paragraph.AppendChild(new Run(doc, " [processed]"));

            // Report progress as a fraction of total work completed.
            progress.Report((i + 1) / (double)total);
        }
    }
}
