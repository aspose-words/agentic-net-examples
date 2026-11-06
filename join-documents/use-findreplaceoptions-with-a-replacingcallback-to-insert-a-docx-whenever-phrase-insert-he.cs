using System;
using System.IO;
using System.Linq;
using Aspose.Words;
using Aspose.Words.BuildingBlocks;

public class Program
{
    public static void Main()
    {
        // ---------- Create destination document with a placeholder ----------
        Document destDoc = new Document();
        DocumentBuilder destBuilder = new DocumentBuilder(destDoc);
        destBuilder.Writeln("This is the main document.");
        destBuilder.Writeln("INSERT_HERE"); // placeholder
        destBuilder.Writeln("End of the main document.");

        // ---------- Create source document that will be inserted ----------
        Document srcDoc = new Document();
        DocumentBuilder srcBuilder = new DocumentBuilder(srcDoc);
        srcBuilder.Writeln("=== Inserted Document Start ===");
        srcBuilder.Writeln("This content comes from the source DOCX.");
        srcBuilder.Writeln("=== Inserted Document End ===");

        // ---------- Locate the placeholder run ----------
        Run? placeholderRun = destDoc.GetChildNodes(NodeType.Run, true)
                                    .Cast<Run>()
                                    .FirstOrDefault(r => r.Text.Contains("INSERT_HERE"));

        if (placeholderRun == null)
            throw new InvalidOperationException("Placeholder text not found in destination document.");

        // ---------- Insert the source document at the placeholder ----------
        DocumentBuilder insertBuilder = new DocumentBuilder(destDoc);
        insertBuilder.MoveTo(placeholderRun);
        insertBuilder.InsertDocument(srcDoc, ImportFormatMode.KeepSourceFormatting);

        // Remove the placeholder run after insertion.
        placeholderRun.Remove();

        // ---------- Save the merged document ----------
        string resultPath = "Result.docx";
        destDoc.Save(resultPath, SaveFormat.Docx);

        // ---------- Validate output ----------
        if (!File.Exists(resultPath))
            throw new InvalidOperationException("Result document was not created.");

        // Optional: clean up (if any temporary files were created, none in this example)
        Console.WriteLine($"Merged document saved to '{resultPath}'.");
    }
}
