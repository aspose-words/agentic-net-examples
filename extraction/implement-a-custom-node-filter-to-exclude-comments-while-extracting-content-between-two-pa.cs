using System;
using System.IO;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // -----------------------------------------------------------------
        // 1. Create a sample source document with a comment.
        // -----------------------------------------------------------------
        Document source = new Document();
        DocumentBuilder builder = new DocumentBuilder(source);

        builder.Writeln("Paragraph 1");
        builder.Writeln("Paragraph 2 with comment");

        // Add a comment anchored to the second paragraph.
        Comment comment = new Comment(source, "Alice", "A", DateTime.Today);
        comment.SetText("This is a comment.");
        builder.CurrentParagraph.AppendChild(comment);

        builder.Writeln("Paragraph 3");
        builder.Writeln("Paragraph 4");

        const string sourcePath = "source.docx";
        source.Save(sourcePath);

        // -----------------------------------------------------------------
        // 2. Load the document for processing.
        // -----------------------------------------------------------------
        Document loaded = new Document(sourcePath);

        // Identify the start and end paragraphs (inclusive range).
        Paragraph startParagraph = loaded.FirstSection.Body.Paragraphs[1]; // "Paragraph 2 with comment"
        Paragraph endParagraph = loaded.FirstSection.Body.Paragraphs[2];   // "Paragraph 3"

        if (startParagraph == null || endParagraph == null)
            throw new InvalidOperationException("Boundary paragraphs not found.");

        // -----------------------------------------------------------------
        // 3. Prepare the result document (empty structure).
        // -----------------------------------------------------------------
        Document result = new Document();
        result.RemoveAllChildren();

        Section resultSection = new Section(result);
        result.AppendChild(resultSection);

        Body resultBody = new Body(result);
        resultSection.AppendChild(resultBody);

        // -----------------------------------------------------------------
        // 4. Determine the indices of the start and end paragraphs.
        // -----------------------------------------------------------------
        int startIndex = loaded.FirstSection.Body.Paragraphs.IndexOf(startParagraph);
        int endIndex = loaded.FirstSection.Body.Paragraphs.IndexOf(endParagraph);

        if (startIndex < 0 || endIndex < 0 || startIndex > endIndex)
            throw new InvalidOperationException("Invalid paragraph range.");

        // -----------------------------------------------------------------
        // 5. Import paragraphs into the result document and strip comments.
        // -----------------------------------------------------------------
        // NodeImporter handles the document ownership transition.
        NodeImporter importer = new NodeImporter(loaded, result, ImportFormatMode.KeepSourceFormatting);

        for (int i = startIndex; i <= endIndex; i++)
        {
            Paragraph srcPara = loaded.FirstSection.Body.Paragraphs[i];

            // Import the paragraph (deep clone) into the destination document.
            Paragraph importedPara = (Paragraph)importer.ImportNode(srcPara, true);

            // Remove any comment-related nodes from the imported paragraph.
            NodeCollection children = importedPara.GetChildNodes(NodeType.Any, true);
            for (int j = children.Count - 1; j >= 0; j--)
            {
                Node child = children[j];
                if (child.NodeType == NodeType.Comment ||
                    child.NodeType == NodeType.CommentRangeStart ||
                    child.NodeType == NodeType.CommentRangeEnd)
                {
                    child.Remove();
                }
            }

            resultBody.AppendChild(importedPara);
        }

        // -----------------------------------------------------------------
        // 6. Save the extracted document and its plain‑text representation.
        // -----------------------------------------------------------------
        const string resultDocPath = "extracted.docx";
        result.Save(resultDocPath);

        string extractedText = result.GetText();
        const string resultTxtPath = "extracted.txt";
        File.WriteAllText(resultTxtPath, extractedText);

        // -----------------------------------------------------------------
        // 7. Validate that the output files were created.
        // -----------------------------------------------------------------
        if (!File.Exists(resultDocPath))
            throw new InvalidOperationException("Extracted document was not created.");

        if (!File.Exists(resultTxtPath))
            throw new InvalidOperationException("Extracted text file was not created.");

        Console.WriteLine("Extraction completed successfully.");
    }
}
