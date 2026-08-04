using System;
using System.IO;
using System.Threading;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Create a sample document with several paragraphs.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        for (int i = 1; i <= 10; i++)
        {
            builder.Writeln($"Paragraph {i}");
        }

        // Save the source document locally.
        string sourcePath = "Source.docx";
        doc.Save(sourcePath);

        // Load the document back from the file system.
        Document loadedDoc = new Document(sourcePath);

        // Prepare a cancellation token source.
        using CancellationTokenSource cts = new CancellationTokenSource();

        // Retrieve all paragraphs in the document.
        NodeCollection paragraphs = loadedDoc.GetChildNodes(NodeType.Paragraph, true);
        int index = 0;

        // Process paragraphs in a while loop, checking for cancellation.
        while (index < paragraphs.Count)
        {
            // Exit gracefully if cancellation is requested.
            if (cts.Token.IsCancellationRequested)
                break;

            Paragraph para = (Paragraph)paragraphs[index];

            // Example processing: append a marker to each paragraph.
            para.AppendChild(new Run(loadedDoc, " - processed"));

            index++;

            // For demonstration, request cancellation after processing five paragraphs.
            if (index == 5)
                cts.Cancel();
        }

        // Save the processed document.
        string outputPath = "Processed.docx";
        loadedDoc.Save(outputPath);

        // Verify that the output file was created.
        if (!File.Exists(outputPath))
            throw new Exception("The processed document was not saved successfully.");
    }
}
