using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Notes;

class Program
{
    static void Main()
    {
        const string inputPath = "Input.docx";
        const string outputPath = "Output.docx";

        // Ensure there is an input document. If it does not exist, create a simple one with a footnote.
        if (!File.Exists(inputPath))
        {
            var tempDoc = new Document();
            var builder = new DocumentBuilder(tempDoc);
            builder.Writeln("This is a sample paragraph with a footnote.");
            builder.InsertFootnote(FootnoteType.Footnote, "This is the footnote text that might be hyphenated.");
            tempDoc.Save(inputPath);
        }

        // Load the original DOCX.
        Document doc = new Document(inputPath);

        // Disable automatic hyphenation for footnote paragraphs only.
        foreach (Footnote footnote in doc.GetChildNodes(NodeType.Footnote, true))
        {
            foreach (Paragraph para in footnote.GetChildNodes(NodeType.Paragraph, true))
            {
                para.ParagraphFormat.SuppressAutoHyphens = true;
            }
        }

        // Save the modified document.
        doc.Save(outputPath);

        Console.WriteLine("Footnote hyphenation disabled and document saved.");
    }
}
