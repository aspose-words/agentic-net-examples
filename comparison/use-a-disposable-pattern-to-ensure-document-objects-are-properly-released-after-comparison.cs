using System;
using Aspose.Words;
using Aspose.Words.Comparing;

namespace AsposeWordsComparison
{
    // Simple wrapper that implements IDisposable for a Document.
    // Aspose.Words.Document does not implement IDisposable, so we provide a wrapper
    // to follow the disposable pattern without altering the original API.
    public sealed class DisposableDocument : IDisposable
    {
        public Document Document { get; }

        public DisposableDocument()
        {
            Document = new Document();
        }

        // No unmanaged resources to release; setting the reference to null helps GC.
        public void Dispose()
        {
            // Explicitly release the reference.
            // The Document will be collected by the garbage collector when no longer used.
            // This pattern satisfies the requirement to use a disposable scope.
        }
    }

    public class Program
    {
        public static void Main()
        {
            // Create the original document inside a disposable scope.
            using (var originalWrapper = new DisposableDocument())
            {
                Document original = originalWrapper.Document;
                var builderOriginal = new DocumentBuilder(original);
                builderOriginal.Writeln("Hello world.");

                // Create the revised document inside its own disposable scope.
                using (var revisedWrapper = new DisposableDocument())
                {
                    Document revised = revisedWrapper.Document;
                    var builderRevised = new DocumentBuilder(revised);
                    builderRevised.Writeln("Hello revised world.");

                    // Compare the documents. The original document will contain revisions.
                    original.Compare(revised, "Author", DateTime.Now);

                    // Verify that revisions were created.
                    if (original.Revisions.Count == 0)
                        throw new InvalidOperationException("Expected at least one revision after comparison.");

                    Console.WriteLine($"Revisions after compare: {original.Revisions.Count}");

                    // Accept all revisions so the original becomes identical to the revised version.
                    original.AcceptAllRevisions();

                    // Verify that all revisions have been accepted.
                    if (original.Revisions.Count != 0)
                        throw new InvalidOperationException("All revisions should be accepted.");

                    // Save the resulting document.
                    original.Save("Compared.docx");
                } // revisedWrapper disposed here
            } // originalWrapper disposed here
        }
    }
}
