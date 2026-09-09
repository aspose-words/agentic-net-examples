using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;
using Aspose.Words.Notes;

public class Program
{
    public static void Main()
    {
        // Create the destination document with a footnote.
        Document destinationDoc = new Document();
        DocumentBuilder destBuilder = new DocumentBuilder(destinationDoc);
        destBuilder.Writeln("Destination document start.");
        destBuilder.InsertFootnote(FootnoteType.Footnote, "Destination footnote 1.");
        destBuilder.Writeln("More text in destination.");

        // Ensure footnote numbering continues (default behavior).
        destinationDoc.FootnoteOptions.RestartRule = FootnoteNumberingRule.Continuous;

        // Create the source document that also contains a footnote.
        Document sourceDoc = new Document();
        DocumentBuilder srcBuilder = new DocumentBuilder(sourceDoc);
        srcBuilder.Writeln("Source document start.");
        srcBuilder.InsertFootnote(FootnoteType.Footnote, "Source footnote 1.");
        srcBuilder.Writeln("More text in source.");

        // Append the source document to the destination document.
        destinationDoc.AppendDocument(sourceDoc, ImportFormatMode.KeepSourceFormatting);

        // Save the merged document as PDF.
        string outputPath = "MergedOutput.pdf";
        destinationDoc.Save(outputPath, SaveFormat.Pdf);

        // Validation: check that the PDF file was created.
        if (!File.Exists(outputPath))
            throw new InvalidOperationException("The merged PDF file was not created.");

        // Validation: ensure the merged document contains both footnotes.
        int footnoteCount = destinationDoc.GetChildNodes(NodeType.Footnote, true).Count;
        if (footnoteCount != 2)
            throw new InvalidOperationException($"Expected 2 footnotes after merge, but found {footnoteCount}.");

        // Optional: indicate successful completion (no console output required).
    }
}
