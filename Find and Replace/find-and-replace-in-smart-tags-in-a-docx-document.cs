using System;
using System.Text.RegularExpressions;
using Aspose.Words;
using Aspose.Words.Markup;
using Aspose.Words.Replacing;

class SmartTagReplacingCallback : IReplacingCallback
{
    // Called for each match found by the Find/Replace engine.
    public ReplaceAction Replacing(ReplacingArgs args)
    {
        // Walk up the node hierarchy to see if the match is inside a SmartTag.
        Node current = args.MatchNode;
        while (current != null && !(current is SmartTag))
            current = current.ParentNode;

        // If the match is within a SmartTag, allow the replacement.
        if (current is SmartTag)
            return ReplaceAction.Replace;

        // Otherwise skip this match – we only want to replace inside SmartTags.
        return ReplaceAction.Skip;
    }
}

class Program
{
    static void Main()
    {
        // Load the DOCX document.
        Document doc = new Document("Input.docx");

        // Configure find/replace options with our custom callback.
        FindReplaceOptions options = new FindReplaceOptions();
        options.ReplacingCallback = new SmartTagReplacingCallback();

        // Example: replace the word "oldValue" with "newValue" only inside SmartTags.
        Regex findPattern = new Regex("oldValue", RegexOptions.IgnoreCase);
        doc.Range.Replace(findPattern, "newValue", options);

        // Save the modified document.
        doc.Save("Output.docx");
    }
}
