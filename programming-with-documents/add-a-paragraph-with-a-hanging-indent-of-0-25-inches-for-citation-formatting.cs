using System;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Sample citation text that may wrap onto multiple lines.
        string citation = "Doe, J. (2020). Example citation text that may wrap onto multiple lines.";

        // Insert the citation as a new paragraph.
        builder.Writeln(citation);

        // Retrieve the paragraph that was just added.
        Paragraph paragraph = (Paragraph)doc.GetChild(NodeType.Paragraph, doc.GetChildNodes(NodeType.Paragraph, true).Count - 1, true);

        // Set a hanging indent of 0.25 inches.
        // LeftIndent = 0.25 inches, FirstLineIndent = -0.25 inches (negative creates hanging indent).
        double indent = ConvertUtil.InchToPoint(0.25);
        paragraph.ParagraphFormat.LeftIndent = indent;
        paragraph.ParagraphFormat.FirstLineIndent = -indent;

        // Save the document to a file.
        string outputPath = "HangingIndent.docx";
        doc.Save(outputPath);
    }
}
