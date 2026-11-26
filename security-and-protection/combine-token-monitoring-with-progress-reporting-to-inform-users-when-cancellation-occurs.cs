using System;
using System.Collections.Generic;
using System.IO;
using System.Threading;
using System.Threading.Tasks;

namespace Aspose.Words
{
    // Minimal stub enums to mimic Aspose.Words types
    public enum ProtectionType { ReadOnly }
    public enum NodeType { Paragraph }

    // Stub for LoadOptions used when opening a protected document
    public class LoadOptions
    {
        public string Password { get; set; }
    }

    // Simple node representation
    public class Node
    {
        public NodeType NodeType { get; set; }
        public string Text { get; set; }
    }

    // Very small Document implementation sufficient for the demo
    public class Document
    {
        private readonly List<Node> _nodes = new List<Node>();
        private bool _isProtected;
        private string _password;

        public Document() { }

        // Load from file using (ignored) LoadOptions – password handling is simulated
        public Document(string filePath, LoadOptions loadOptions)
        {
            if (!File.Exists(filePath))
                throw new FileNotFoundException("Document not found.", filePath);

            // In this stub we simply read all lines as separate paragraph nodes.
            foreach (var line in File.ReadAllLines(filePath))
            {
                _nodes.Add(new Node { NodeType = NodeType.Paragraph, Text = line });
            }

            // Simulate password check – if the document was saved as protected we expect a password.
            // For the stub we just store the password that was used when protecting.
            // Real Aspose.Words would validate the password here.
            // No action needed for the demo.
        }

        // Save the document – writes each paragraph as a line.
        public void Save(string filePath)
        {
            var lines = new List<string>();
            foreach (var node in _nodes)
            {
                if (node.NodeType == NodeType.Paragraph)
                    lines.Add(node.Text);
            }
            File.WriteAllLines(filePath, lines);
        }

        // Apply simple read‑only protection.
        public void Protect(ProtectionType protectionType, string password)
        {
            if (protectionType != ProtectionType.ReadOnly)
                throw new NotSupportedException("Only ReadOnly protection is supported in this stub.");

            _isProtected = true;
            _password = password;
        }

        // Retrieve child nodes of a given type.
        public List<Node> GetChildNodes(NodeType nodeType, bool deep)
        {
            // In this stub we ignore the 'deep' flag because the document structure is flat.
            return _nodes.FindAll(n => n.NodeType == nodeType);
        }
    }

    // Simple builder that adds paragraphs to a Document.
    public class DocumentBuilder
    {
        private readonly Document _doc;

        public DocumentBuilder(Document doc)
        {
            _doc = doc;
        }

        public void Writeln(string text)
        {
            // Add a new paragraph node.
            var nodes = _doc.GetChildNodes(NodeType.Paragraph, true);
            // Directly manipulate the internal list via reflection (since we keep it simple).
            var field = typeof(Document).GetField("_nodes", System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Instance);
            var list = (List<Node>)field.GetValue(_doc);
            list.Add(new Node { NodeType = NodeType.Paragraph, Text = text });
        }
    }
}

public class Program
{
    public static void Main()
    {
        // Prepare file paths
        string tempDir = Path.Combine(Path.GetTempPath(), "AsposeDemo");
        Directory.CreateDirectory(tempDir);
        string originalPath = Path.Combine(tempDir, "original.docx");
        string protectedPath = Path.Combine(tempDir, "protected.docx");

        // 1. Create a simple document
        var doc = new Aspose.Words.Document();
        var builder = new Aspose.Words.DocumentBuilder(doc);
        builder.Writeln("Hello world! This is a sample document for processing.");
        builder.Writeln("Second paragraph.");
        builder.Writeln("Third paragraph.");
        doc.Save(originalPath);

        // 2. Apply read‑only protection with a password and save
        const string password = "myPassword";
        doc.Protect(Aspose.Words.ProtectionType.ReadOnly, password);
        doc.Save(protectedPath);

        // Validate that the protected file exists
        if (!File.Exists(protectedPath))
            throw new FileNotFoundException("Protected document was not saved.", protectedPath);

        // 3. Load the protected document using the password
        var loadOptions = new Aspose.Words.LoadOptions { Password = password };
        var protectedDoc = new Aspose.Words.Document(protectedPath, loadOptions);

        // 4. Set up progress reporting and cancellation
        var progress = new Progress<int>(percent =>
            Console.WriteLine($"Processing progress: {percent}%"));
        using var cts = new CancellationTokenSource();

        // Cancel after 350 ms to simulate user cancellation
        Task.Delay(350).ContinueWith(_ => cts.Cancel());

        try
        {
            ProcessDocumentAsync(protectedDoc, progress, cts.Token).GetAwaiter().GetResult();
            Console.WriteLine("Processing completed successfully.");
        }
        catch (OperationCanceledException)
        {
            Console.WriteLine("Processing was cancelled by the token.");
        }
        finally
        {
            // Clean up temporary files
            try { File.Delete(originalPath); } catch { }
            try { File.Delete(protectedPath); } catch { }
            try { Directory.Delete(tempDir, true); } catch { }
        }
    }

    private static async Task ProcessDocumentAsync(Aspose.Words.Document document, IProgress<int> progress, CancellationToken token)
    {
        // Simulate processing each paragraph with delay
        var paragraphs = document.GetChildNodes(Aspose.Words.NodeType.Paragraph, true);
        int total = paragraphs.Count;
        for (int i = 0; i < total; i++)
        {
            token.ThrowIfCancellationRequested();

            // Simulated work (e.g., analyzing paragraph text)
            await Task.Delay(100, token);

            int percent = (i + 1) * 100 / total;
            progress.Report(percent);
        }
    }
}
