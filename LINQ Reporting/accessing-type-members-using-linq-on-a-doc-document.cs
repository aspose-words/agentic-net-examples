using System;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Fields;
using Aspose.Words.Properties;

class Program
{
    static void Main()
    {
        // Load an existing DOC/DOCX document.
        Document doc = new Document("Input.docx");

        // -------------------------------------------------
        // Example 1: Use LINQ to find all paragraphs that contain the word "Aspose".
        var paragraphsWithAspose = doc.GetChildNodes(NodeType.Paragraph, true)
                                      .OfType<Paragraph>()
                                      .Where(p => p.GetText().Contains("Aspose"));

        foreach (var para in paragraphsWithAspose)
        {
            Console.WriteLine("Paragraph: " + para.GetText().Trim());
        }

        // -------------------------------------------------
        // Example 2: Use LINQ to retrieve all hyperlink fields in the document.
        var hyperlinkFields = doc.SelectNodes("//FieldStart")
                                 .OfType<FieldStart>()
                                 .Where(fs => fs.FieldType == FieldType.FieldHyperlink);

        foreach (var fieldStart in hyperlinkFields)
        {
            // The Hyperlink helper class from the documentation parses the field.
            var hyperlink = new Hyperlink(fieldStart);
            Console.WriteLine($"Hyperlink Target: {hyperlink.Target}, Name: {hyperlink.Name}");
        }

        // -------------------------------------------------
        // Example 3: Use LINQ to list all custom document properties of type string.
        var stringProperties = doc.CustomDocumentProperties
                                  .Cast<DocumentProperty>()
                                  .Where(p => p.Type == PropertyType.String)
                                  .Select(p => new { p.Name, Value = p.Value.ToString() });

        foreach (var prop in stringProperties)
        {
            Console.WriteLine($"Custom Property - Name: {prop.Name}, Value: {prop.Value}");
        }

        // -------------------------------------------------
        // Example 4: Use LINQ to find all runs that are bold.
        var boldRuns = doc.GetChildNodes(NodeType.Run, true)
                          .OfType<Run>()
                          .Where(r => r.Font.Bold);

        foreach (var run in boldRuns)
        {
            Console.WriteLine("Bold Run Text: " + run.Text);
        }

        // Save the modified document (if any changes were made).
        doc.Save("Output.docx");
    }

    // Helper class to work with hyperlink fields (simplified version from Aspose.Words examples).
    private class Hyperlink
    {
        private readonly FieldStart _fieldStart;
        private readonly Node _fieldSeparator;
        private readonly Node _fieldEnd;
        private string _target;
        private bool _isLocal;

        public Hyperlink(FieldStart fieldStart)
        {
            if (fieldStart == null) throw new ArgumentNullException(nameof(fieldStart));
            if (fieldStart.FieldType != FieldType.FieldHyperlink)
                throw new ArgumentException("Field start type must be FieldHyperlink.");

            _fieldStart = fieldStart;
            _fieldSeparator = FindNextSibling(_fieldStart, NodeType.FieldSeparator);
            _fieldEnd = FindNextSibling(_fieldSeparator, NodeType.FieldEnd);

            string fieldCode = GetTextSameParent(_fieldStart.NextSibling, _fieldSeparator);
            var match = System.Text.RegularExpressions.Regex.Match(fieldCode.Trim(),
                @"\\l\s+|""([^""]+)""");
            _isLocal = match.Value.StartsWith("\\l");
            _target = match.Groups[1].Value;
        }

        public string Target => _target;
        public string Name => GetTextSameParent(_fieldSeparator, _fieldEnd);

        private static Node FindNextSibling(Node startNode, NodeType nodeType)
        {
            for (Node node = startNode; node != null; node = node.NextSibling)
                if (node.NodeType == nodeType) return node;
            return null;
        }

        private static string GetTextSameParent(Node startNode, Node endNode)
        {
            if (endNode != null && startNode.ParentNode != endNode.ParentNode)
                throw new ArgumentException("Start and end nodes must have the same parent.");

            var builder = new System.Text.StringBuilder();
            for (Node cur = startNode; cur != null && cur != endNode; cur = cur.NextSibling)
                builder.Append(cur.GetText());
            return builder.ToString();
        }
    }
}
