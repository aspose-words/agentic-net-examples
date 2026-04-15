using System;
using System.Collections.Generic;
using Aspose.Words;

public class SplitParagraphExample
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Add a single long paragraph.
        string longText = "Lorem ipsum dolor sit amet, consectetur adipiscing elit. " +
                          "Sed non risus sit amet elit placerat gravida. " +
                          "Praesent euismod, nisi a interdum consequat, " +
                          "nisi lorem fermentum odio, a bibendum sapien " +
                          "quam a justo. Integer nec odio nec urna " +
                          "vehicula tincidunt. Donec vitae ligula " +
                          "vitae sapien ultricies aliquet. Curabitur " +
                          "sagittis, massa at fermentum commodo, " +
                          "nunc elit lacinia odio, a interdum " +
                          "nunc sapien non justo.";
        builder.Writeln(longText);

        // Define split positions (character indices) where the paragraph will be broken.
        int[] splitPositions = { 100, 200, 300 };

        // Retrieve the original paragraph and its text (without the trailing paragraph break character).
        Paragraph originalParagraph = doc.FirstSection.Body.FirstParagraph;
        string originalText = originalParagraph.GetText().TrimEnd('\r');

        // Split the text according to the specified positions.
        List<string> parts = new List<string>();
        int start = 0;
        foreach (int pos in splitPositions)
        {
            int end = Math.Min(pos, originalText.Length);
            if (end > start)
            {
                parts.Add(originalText.Substring(start, end - start));
                start = end;
            }
        }
        // Add the remaining text after the last split position.
        if (start < originalText.Length)
            parts.Add(originalText.Substring(start));

        // Remove the original paragraph from the document.
        originalParagraph.Remove();

        // Move the builder to the start of the document body so that new paragraphs are inserted correctly.
        builder.MoveToDocumentStart();

        // Insert the new shorter paragraphs.
        foreach (string part in parts)
        {
            builder.Writeln(part);
        }

        // Save the resulting document.
        doc.Save("SplitParagraph.docx");
    }
}
