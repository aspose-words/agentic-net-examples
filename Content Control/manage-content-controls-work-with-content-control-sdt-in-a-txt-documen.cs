using System;
using Aspose.Words;
using Aspose.Words.Markup;

class ContentControlDemo
{
    static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();

        // Initialize a DocumentBuilder for the document.
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert a plain‑text content control (SDT) at the current cursor position.
        StructuredDocumentTag sdt = new StructuredDocumentTag(doc, SdtType.PlainText, MarkupLevel.Inline);

        // Set some useful properties of the content control.
        sdt.Title = "CustomerName";                     // Friendly name.
        sdt.Tag = "CustomerNameTag";                    // Tag identifier.
        sdt.LockContents = true;                         // Prevent editing of the content.
        sdt.IsShowingPlaceholderText = true;             // Show placeholder when empty.

        // Insert the content control into the document.
        builder.InsertNode(sdt);

        // Move the cursor inside the content control to add placeholder text.
        // StructuredDocumentTag does not expose a FirstParagraph property; use MoveTo(sdt) instead.
        builder.MoveTo(sdt);
        builder.Write("Enter name here...");

        // Save the document to a file (format inferred from extension).
        doc.Save("ContentControlDemo.docx");

        // -----------------------------------------------------------------
        // Load the saved document and demonstrate how to read the content control.
        // -----------------------------------------------------------------
        Document loadedDoc = new Document("ContentControlDemo.docx");
        StructuredDocumentTag loadedSdt = (StructuredDocumentTag)loadedDoc.GetChildNodes(NodeType.StructuredDocumentTag, true)[0];

        // Output some properties to the console.
        Console.WriteLine("Title: " + loadedSdt.Title);
        Console.WriteLine("Tag: " + loadedSdt.Tag);
        Console.WriteLine("Locked: " + loadedSdt.LockContents);
        Console.WriteLine("Placeholder shown: " + loadedSdt.IsShowingPlaceholderText);
        Console.WriteLine("Current text: " + loadedSdt.GetText().Trim());
    }
}
