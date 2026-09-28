using System;
using System.IO;
using System.Threading.Tasks;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Create a sample document with many paragraphs to make layout generation take noticeable time.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        for (int i = 0; i < 2000; i++)
        {
            builder.Writeln($"Paragraph {i + 1}: The quick brown fox jumps over the lazy dog.");
        }

        // Save the source document (required by the rules to persist the file).
        string sourcePath = "SourceDocument.docx";
        doc.Save(sourcePath);
        if (!File.Exists(sourcePath))
            throw new Exception("Failed to create the source document.");

        // Start layout generation in a separate task.
        Task layoutTask = Task.Run(() => doc.UpdatePageLayout());

        // Simulate a user abort by waiting only a short time for the layout to finish.
        bool layoutCompleted = layoutTask.Wait(TimeSpan.FromMilliseconds(100));

        if (layoutCompleted)
        {
            Console.WriteLine("Layout completed without cancellation.");
        }
        else
        {
            Console.WriteLine("Layout was cancelled (timeout) as expected.");
            // Optionally, you could ignore the incomplete layout or take other actions here.
        }

        // Save the document after attempting layout (if it wasn't cancelled, it will be saved).
        string outputPath = "LayoutResult.docx";
        doc.Save(outputPath);
        if (!File.Exists(outputPath))
            throw new Exception("Failed to save the output document.");

        // Indicate successful execution.
        Console.WriteLine("Execution finished.");
    }
}
