using System;
using System.IO;
using Aspose.Words;

public class Program
{
    // Simple wrapper to apply the disposable pattern to Aspose.Words.Document.
    private sealed class DisposableDocument : IDisposable
    {
        public Document Doc { get; }

        public DisposableDocument()
        {
            Doc = new Document(); // Create a blank document.
        }

        public void Dispose()
        {
            // Document does not hold unmanaged resources that require explicit disposal,
            // but implementing IDisposable allows a consistent using‑statement pattern.
        }
    }

    public static void Main()
    {
        // Create the original document with some content.
        using (var originalWrapper = new DisposableDocument())
        {
            DocumentBuilder builder1 = new DocumentBuilder(originalWrapper.Doc);
            builder1.Writeln("Hello world.");
            
            // Create the revised document that differs from the original.
            using (var revisedWrapper = new DisposableDocument())
            {
                DocumentBuilder builder2 = new DocumentBuilder(revisedWrapper.Doc);
                builder2.Writeln("Hello revised world.");

                // Perform the comparison. The original document will receive revisions.
                originalWrapper.Doc.Compare(revisedWrapper.Doc, "JD", DateTime.Now);

                // Verify that at least one revision was created.
                if (originalWrapper.Doc.Revisions.Count == 0)
                    throw new InvalidOperationException("Expected at least one revision after comparison.");

                // Accept all revisions so the original document becomes identical to the revised one.
                originalWrapper.Doc.AcceptAllRevisions();

                // Ensure that all revisions have been cleared.
                if (originalWrapper.Doc.Revisions.Count != 0)
                    throw new InvalidOperationException("All revisions should have been accepted.");

                // Save the final document to the local file system.
                originalWrapper.Doc.Save("ComparisonResult.docx");
            }
        }
    }
}
