using System;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        try
        {
            VerifyParagraphRangeExtractionPreservesStyling();
            Console.WriteLine("Test passed.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Test failed: {ex.Message}");
        }
    }

    private static void VerifyParagraphRangeExtractionPreservesStyling()
    {
        // ---------- Create source document with styled paragraphs ----------
        Document sourceDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(sourceDoc);

        // Paragraph before the range (no special styling)
        builder.Writeln("Paragraph before start.");

        // Start paragraph – bold text
        builder.Font.Bold = true;
        builder.Writeln("Start paragraph - bold.");

        // Middle paragraph – italic text
        builder.Font.Bold = false;
        builder.Font.Italic = true;
        builder.Writeln("Middle paragraph - italic.");

        // End paragraph – underline text
        builder.Font.Italic = false;
        builder.Font.Underline = Underline.Single;
        builder.Writeln("End paragraph - underline.");

        // Paragraph after the range (no special styling)
        builder.Font.Underline = Underline.None;
        builder.Writeln("Paragraph after end.");

        // Save source for manual inspection (optional)
        sourceDoc.Save("Source.docx");

        // ---------- Locate start and end paragraphs ----------
        NodeCollection allParagraphs = sourceDoc.GetChildNodes(NodeType.Paragraph, true);
        // Indices are zero‑based; we know the order from the builder calls.
        int startIndex = 1; // "Start paragraph - bold."
        int endIndex = 3;   // "End paragraph - underline."

        if (startIndex < 0 || endIndex >= allParagraphs.Count || startIndex > endIndex)
            throw new InvalidOperationException("Invalid paragraph range indices.");

        // ---------- Prepare destination document ----------
        Document destDoc = new Document();
        destDoc.RemoveAllChildren(); // Remove the default empty section/paragraph.
        Section destSection = new Section(destDoc);
        destDoc.AppendChild(destSection);
        Body destBody = new Body(destDoc);
        destSection.AppendChild(destBody);

        // ---------- Import the selected paragraph range ----------
        NodeImporter importer = new NodeImporter(sourceDoc, destDoc, ImportFormatMode.KeepSourceFormatting);
        for (int i = startIndex; i <= endIndex; i++)
        {
            Paragraph srcParagraph = (Paragraph)allParagraphs[i];
            Node importedNode = importer.ImportNode(srcParagraph, true);
            destBody.AppendChild(importedNode);
        }

        // Save extracted document for manual inspection (optional)
        destDoc.Save("Extracted.docx");

        // ---------- Validation ----------
        // 1. Paragraph count must match.
        int expectedCount = endIndex - startIndex + 1;
        if (destBody.Paragraphs.Count != expectedCount)
            throw new Exception($"Extracted paragraph count {destBody.Paragraphs.Count} does not match expected {expectedCount}.");

        // 2. Verify that each run's font styling is preserved.
        for (int i = 0; i < expectedCount; i++)
        {
            Paragraph srcPara = (Paragraph)allParagraphs[startIndex + i];
            Paragraph dstPara = destBody.Paragraphs[i];

            // Compare each run in the paragraph.
            NodeCollection srcRuns = srcPara.GetChildNodes(NodeType.Run, true);
            NodeCollection dstRuns = dstPara.GetChildNodes(NodeType.Run, true);

            if (srcRuns.Count != dstRuns.Count)
                throw new Exception($"Run count mismatch in paragraph {i}.");

            for (int r = 0; r < srcRuns.Count; r++)
            {
                Run srcRun = (Run)srcRuns[r];
                Run dstRun = (Run)dstRuns[r];

                if (srcRun.Font.Bold != dstRun.Font.Bold ||
                    srcRun.Font.Italic != dstRun.Font.Italic ||
                    srcRun.Font.Underline != dstRun.Font.Underline)
                {
                    throw new Exception($"Font styling mismatch in paragraph {i}, run {r}.");
                }
            }
        }
    }
}
