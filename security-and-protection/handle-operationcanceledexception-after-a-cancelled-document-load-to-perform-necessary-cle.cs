using System;
using System.IO;
using System.Threading;
using Aspose.Words;
using Aspose.Words.Loading;

public class Program
{
    // Stream wrapper that throws OperationCanceledException when the token is cancelled.
    private class CancellableStream : Stream
    {
        private readonly Stream _baseStream;
        private readonly CancellationToken _token;

        public CancellableStream(Stream baseStream, CancellationToken token)
        {
            _baseStream = baseStream;
            _token = token;
        }

        public override bool CanRead => _baseStream.CanRead;
        public override bool CanSeek => _baseStream.CanSeek;
        public override bool CanWrite => _baseStream.CanWrite;
        public override long Length => _baseStream.Length;

        public override long Position
        {
            get => _baseStream.Position;
            set => _baseStream.Position = value;
        }

        public override void Flush() => _baseStream.Flush();

        public override int Read(byte[] buffer, int offset, int count)
        {
            _token.ThrowIfCancellationRequested();
            return _baseStream.Read(buffer, offset, count);
        }

        public override long Seek(long offset, SeekOrigin origin) => _baseStream.Seek(offset, origin);
        public override void SetLength(long value) => _baseStream.SetLength(value);
        public override void Write(byte[] buffer, int offset, int count) => _baseStream.Write(buffer, offset, count);

        protected override void Dispose(bool disposing)
        {
            if (disposing)
                _baseStream.Dispose();
            base.Dispose(disposing);
        }
    }

    public static void Main()
    {
        // Create a simple document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("Sample text for cancellation test.");

        // Save the document to a temporary file.
        string tempPath = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString() + ".docx");
        doc.Save(tempPath);

        // Load the file bytes into a memory stream.
        byte[] fileBytes = File.ReadAllBytes(tempPath);
        using (MemoryStream memoryStream = new MemoryStream(fileBytes))
        {
            // Prepare a cancelled token.
            CancellationTokenSource cts = new CancellationTokenSource();
            cts.Cancel(); // Cancel immediately.

            // Wrap the memory stream with the cancellable stream.
            using (CancellableStream cancellableStream = new CancellableStream(memoryStream, cts.Token))
            {
                try
                {
                    // Attempt to load the document; this should throw OperationCanceledException.
                    Document loadedDoc = new Document(cancellableStream, new LoadOptions());
                    // If load succeeds (unlikely), we simply let it go out of scope.
                }
                catch (OperationCanceledException)
                {
                    // Expected cancellation handling.
                    // Perform any necessary cleanup here.
                    // For this example, we just acknowledge the cancellation.
                    Console.WriteLine("Document loading was cancelled as expected.");
                }
                finally
                {
                    // Ensure the temporary file is deleted regardless of outcome.
                    if (File.Exists(tempPath))
                    {
                        try
                        {
                            File.Delete(tempPath);
                        }
                        catch
                        {
                            // Ignored – cleanup best effort.
                        }
                    }
                }
            }
        }
    }
}
