using System;
using Aspose.Words;
using Aspose.Words.Markup;

public class Program
{
    public static void Main()
    {
        // Step 1: Create a sample document with a bookmark.
        Document seedDoc = new Document();
        DocumentBuilder seedBuilder = new DocumentBuilder(seedDoc);

        seedBuilder.Writeln("Document before the bookmark.");
        seedBuilder.StartBookmark("TargetBookmark");
        seedBuilder.Writeln("Placeholder text inside the bookmark.");
        seedBuilder.EndBookmark("TargetBookmark");
        seedBuilder.Writeln("Document after the bookmark.");

        // Save the seed document.
        const string inputPath = "input.docx";
        seedDoc.Save(inputPath);

        // Step 2: Load the document that contains the bookmark.
        Document doc = new Document(inputPath);
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Move the cursor to the start of the bookmark.
        builder.MoveToBookmark("TargetBookmark");

        // Insert a plain‑text content control at the bookmark position.
        StructuredDocumentTag sdt = builder.InsertStructuredDocumentTag(SdtType.PlainText);
        sdt.Title = "CustomerName";
        sdt.Tag = "customer-name";

        // Replace any existing placeholder text with the desired default content.
        sdt.RemoveAllChildren();
        sdt.AppendChild(new Run(doc, "Contoso"));

        // Save the modified document.
        const string outputPath = "output.docx";
        doc.Save(outputPath);
    }
}
