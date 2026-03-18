using System;
using System.Text.RegularExpressions;
using Aspose.Words;
using Aspose.Words.Fields;
using Aspose.Words.Replacing;

class SectionLetterConverter
{
    static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Add some sample paragraphs containing the placeholder tag.
        builder.Writeln("Introduction");
        builder.Writeln("{=Section:letter} This is the first custom section.");
        builder.Writeln("{=Section:letter} This is the second custom section.");
        builder.Writeln("{=Section:letter} This is the third custom section.");

        // Define a regular expression that matches the placeholder tag.
        Regex placeholderRegex = new Regex(@"\{=Section:letter\}");

        // Replace each occurrence of the placeholder with a SECTION field
        // formatted as an uppercase alphabetic letter (A, B, C, ...).
        FindReplaceOptions options = new FindReplaceOptions();
        options.ReplacingCallback = new SectionLetterCallback();

        doc.Range.Replace(placeholderRegex, string.Empty, options);

        // Save the resulting document.
        doc.Save("SectionLetterFormatted.docx");
    }

    // Callback that replaces the matched placeholder with a SECTION field
    // and applies the UppercaseAlphabetic general format.
    private class SectionLetterCallback : IReplacingCallback
    {
        ReplaceAction IReplacingCallback.Replacing(ReplacingArgs args)
        {
            // The node that contains the matched text.
            Node matchNode = args.MatchNode;

            // Insert the SECTION field at the position of the match.
            // Use the DocumentBuilder to simplify insertion.
            DocumentBuilder builder = new DocumentBuilder((Document)matchNode.Document);
            builder.MoveTo(matchNode);
            FieldSection sectionField = (FieldSection)builder.InsertField(FieldType.FieldSection, true);

            // Apply the UppercaseAlphabetic format to the field result.
            // This converts the numeric section number (1,2,3,…) to A,B,C,…
            sectionField.Format.GeneralFormats.Add(GeneralFormat.UppercaseAlphabetic);

            // Remove the original placeholder text.
            // The match node is a Run; replace its text with an empty string.
            if (matchNode is Run run)
                run.Text = string.Empty;

            // Indicate that the replacement has been handled.
            return ReplaceAction.Skip;
        }
    }
}
