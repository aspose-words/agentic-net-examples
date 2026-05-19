using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Comparing;

namespace ComparisonExample
{
    // Simple wrapper to allow using‑statement disposal pattern for Document.
    // Aspose.Words.Document does not implement IDisposable, so we encapsulate it.
    internal sealed class DocumentHolder : IDisposable
    {
        public Document Document { get; }

        public DocumentHolder(Document document) => Document = document ?? throw new ArgumentNullException(nameof(document));

        public void Dispose()
        {
            // No unmanaged resources in Document, but clearing the reference helps GC.
            // If future versions implement IDisposable, this pattern will still work.
        }
    }

    public class Program
    {
        public static void Main()
        {
            // Prepare a folder for output files.
            string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "ComparisonOutput");
            Directory.CreateDirectory(outputDir);

            // Create the original document inside a using block via the wrapper.
            using (var originalHolder = new DocumentHolder(new Document()))
            {
                Document original = originalHolder.Document;
                var builderOriginal = new DocumentBuilder(original);
                builderOriginal.Writeln("This is the original paragraph.");
                builderOriginal.Writeln("It contains some text.");

                // Create the revised document, also within a using block.
                using (var revisedHolder = new DocumentHolder(new Document()))
                {
                    Document revised = revisedHolder.Document;
                    var builderRevised = new DocumentBuilder(revised);
                    builderRevised.Writeln("This is the original paragraph."); // unchanged line.
                    builderRevised.Writeln("It contains some modified text."); // changed line.

                    // Perform the comparison. The original document will receive revisions.
                    original.Compare(revised, "Comparer", DateTime.Now);

                    // Verify that revisions were generated.
                    if (original.Revisions.Count == 0)
                        throw new InvalidOperationException("Expected at least one revision after comparison.");

                    // Output revision details to the console.
                    foreach (Revision rev in original.Revisions)
                    {
                        Console.WriteLine($"Revision type: {rev.RevisionType}, Author: {rev.Author}");
                        Console.WriteLine($"Changed text: \"{rev.ParentNode.GetText().Trim()}\"");
                    }

                    // Accept all revisions to transform the original into the revised version.
                    original.AcceptAllRevisions();

                    // Verify that all revisions have been accepted.
                    if (original.Revisions.Count != 0)
                        throw new InvalidOperationException("All revisions should be accepted.");

                    // Save the final document.
                    string resultPath = Path.Combine(outputDir, "ComparedResult.docx");
                    original.Save(resultPath);
                    Console.WriteLine($"Comparison result saved to: {resultPath}");
                }
            }
        }
    }
}
