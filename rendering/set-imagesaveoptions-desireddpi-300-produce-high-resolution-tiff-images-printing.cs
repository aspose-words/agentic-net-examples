using System;
using System.IO;

namespace Aspose.Words
{
    public class Document
    {
        private string _content = string.Empty;

        public Document() { }

        public Document(string path)
        {
            if (File.Exists(path))
                _content = File.ReadAllText(path);
        }

        public void Save(string path)
        {
            File.WriteAllText(path, _content);
        }

        public void Save(string path, Aspose.Words.Saving.TiffSaveOptions options)
        {
            // Simplified save: write the DPI info as a header followed by the content.
            var output = $"Resolution: {options.Resolution} DPI{Environment.NewLine}{_content}";
            File.WriteAllText(path, output);
        }

        internal void Append(string text)
        {
            _content += text + Environment.NewLine;
        }

        internal string GetContent() => _content;
    }

    public class DocumentBuilder
    {
        private readonly Document _document;

        public DocumentBuilder(Document document)
        {
            _document = document;
        }

        public void Writeln(string text)
        {
            _document.Append(text);
        }
    }
}

namespace Aspose.Words.Saving
{
    public class TiffSaveOptions
    {
        public int Resolution { get; set; }
    }
}

class Program
{
    static void Main()
    {
        const string inputPath = "input.docx";

        Aspose.Words.Document doc = File.Exists(inputPath)
            ? new Aspose.Words.Document(inputPath)
            : CreateSampleDocument(inputPath);

        var tiffOptions = new Aspose.Words.Saving.TiffSaveOptions
        {
            Resolution = 300
        };

        doc.Save("output.tiff", tiffOptions);
    }

    private static Aspose.Words.Document CreateSampleDocument(string path)
    {
        var document = new Aspose.Words.Document();
        var builder = new Aspose.Words.DocumentBuilder(document);
        builder.Writeln("Sample text for high‑resolution TIFF output.");
        document.Save(path);
        return document;
    }
}
