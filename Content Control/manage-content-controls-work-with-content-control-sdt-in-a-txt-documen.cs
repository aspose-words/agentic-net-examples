using System;
using Aspose.Words;
using Aspose.Words.Markup;
using Aspose.Words.Loading;
using Aspose.Words.Saving;

class ContentControlExample
{
    static void Main()
    {
        // Paths to the input TXT file and the output document.
        string inputPath = "input.txt";
        string outputPath = "output.docx";

        // Load the TXT file using TxtLoadOptions (default options are sufficient here).
        TxtLoadOptions loadOptions = new TxtLoadOptions();
        Document doc = new Document(inputPath, loadOptions);

        // Create a DocumentBuilder to modify the document.
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Move the cursor to the end of the document where we will insert the content control.
        builder.MoveToDocumentEnd();

        // Create a plain‑text structured document tag (content control).
        StructuredDocumentTag sdt = new StructuredDocumentTag(doc, SdtType.PlainText, MarkupLevel.Inline);

        // Insert the content control into the document.
        builder.InsertNode(sdt);

        // Position the builder inside the newly created content control.
        builder.MoveTo(sdt);

        // Add placeholder text inside the content control.
        builder.Writeln("Enter text here...");

        // Set some useful properties of the content control.
        sdt.Title = "MyPlainTextControl";
        sdt.Tag = "PlainTextTag";
        sdt.LockContents = false;          // Allow editing of the contents.
        sdt.LockContentControl = false;    // Allow deletion of the control.

        // Example: retrieve all content controls in the document and print their titles.
        foreach (IStructuredDocumentTag tag in doc.Range.StructuredDocumentTags)
        {
            Console.WriteLine($"SDT Id={tag.Id}, Title={tag.Title}, Tag={tag.Tag}, Type={tag.SdtType}");
        }

        // Save the modified document. The format is inferred from the file extension (DOCX).
        doc.Save(outputPath, SaveFormat.Docx);
    }
}
