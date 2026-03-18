using System;
using System.Text.RegularExpressions;
using Aspose.Words;
using Aspose.Words.Replacing;

class Program
{
    static void Main()
    {
        // Create a new document and a builder attached to it.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Sample content with placeholders that will be replaced.
        builder.Writeln("Dear _Name_,");
        builder.Writeln("Your order _OrderId_ has been shipped.");
        builder.Writeln("Thank you, _Name_.");

        // Set up FindReplaceOptions with a custom callback.
        FindReplaceOptions options = new FindReplaceOptions();
        options.ReplacingCallback = new InsertAfterCallback(builder);

        // Replace the placeholder "_Name_" with an actual name.
        doc.Range.Replace("_Name_", "John Doe", options);
        // Replace the placeholder "_OrderId_" with a dynamic order number.
        doc.Range.Replace("_OrderId_", "12345", options);

        // Save the resulting document.
        doc.Save("DynamicInsertAfterReplace.docx");
    }
}

// Callback that inserts a new paragraph after each successful replacement.
class InsertAfterCallback : IReplacingCallback
{
    private readonly DocumentBuilder _builder;

    public InsertAfterCallback(DocumentBuilder builder)
    {
        _builder = builder;
    }

    public ReplaceAction Replacing(ReplacingArgs args)
    {
        // Perform the normal replacement.
        // (The replacement text is already set in args.Replacement by the caller.)

        // Locate the paragraph that contains the start of the match.
        Paragraph paragraph = args.MatchNode.ParentNode as Paragraph;
        if (paragraph != null)
        {
            // Move the builder to the found paragraph.
            _builder.MoveTo(paragraph);
            // Insert a new paragraph after it with dynamic content.
            _builder.Writeln($"[Inserted after replacement at {DateTime.Now:yyyy-MM-dd HH:mm:ss}]");
        }

        // Continue with the replacement.
        return ReplaceAction.Replace;
    }
}
