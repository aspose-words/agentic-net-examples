using System;
using System.IO;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Math;
using Aspose.Words.Fields;
using Aspose.Words.Saving;

public class ReplaceOfficeMathByBookmark
{
    public static void Main()
    {
        // 1. Create a sample DOCX with two equations, each wrapped in a bookmark.
        Document sampleDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(sampleDoc);

        // First equation with bookmark "eq1"
        builder.StartBookmark("eq1");
        InsertEquation(builder); // creates a deterministic OfficeMath equation
        builder.EndBookmark("eq1");

        // Add a blank paragraph between equations for clarity.
        builder.Writeln();

        // Second equation with bookmark "eq2"
        builder.StartBookmark("eq2");
        InsertEquation(builder);
        builder.EndBookmark("eq2");

        // 2. Save the sample document to a memory stream and reload it to simulate an existing file.
        using (MemoryStream stream = new MemoryStream())
        {
            sampleDoc.Save(stream, SaveFormat.Docx);
            stream.Position = 0;
            Document doc = new Document(stream);

            // 3. Resolve the target bookmark (eq1) and locate its top‑level OfficeMath node.
            Bookmark targetBookmark = doc.Range.Bookmarks["eq1"];
            if (targetBookmark == null)
                throw new InvalidOperationException("Bookmark 'eq1' not found.");

            Paragraph targetParagraph = FindContainingParagraph(targetBookmark.BookmarkStart);
            OfficeMath targetMath = FindTopLevelOfficeMath(targetParagraph);
            if (targetMath == null)
                throw new InvalidOperationException("Target OfficeMath not found in paragraph of 'eq1'.");

            // 4. Resolve the replacement bookmark (eq2) and clone its OfficeMath node.
            Bookmark sourceBookmark = doc.Range.Bookmarks["eq2"];
            if (sourceBookmark == null)
                throw new InvalidOperationException("Bookmark 'eq2' not found.");

            Paragraph sourceParagraph = FindContainingParagraph(sourceBookmark.BookmarkStart);
            OfficeMath sourceMath = FindTopLevelOfficeMath(sourceParagraph);
            if (sourceMath == null)
                throw new InvalidOperationException("Source OfficeMath not found in paragraph of 'eq2'.");

            OfficeMath replacementMath = (OfficeMath)sourceMath.Clone(true);

            // 5. Replace the target equation with the cloned one.
            CompositeNode parent = targetMath.ParentNode as CompositeNode;
            if (parent == null)
                throw new InvalidOperationException("Target OfficeMath does not have a valid parent.");

            parent.InsertBefore(replacementMath, targetMath);
            targetMath.Remove();

            // 6. Save the modified document.
            const string outputPath = "Output.docx";
            doc.Save(outputPath, SaveFormat.Docx);

            // 7. Validation: ensure the file exists and the bookmark still resolves to a paragraph with one top‑level equation.
            if (!File.Exists(outputPath))
                throw new FileNotFoundException("The output file was not created.", outputPath);

            Document finalDoc = new Document(outputPath);
            Bookmark finalBookmark = finalDoc.Range.Bookmarks["eq1"];
            if (finalBookmark == null)
                throw new InvalidOperationException("Bookmark 'eq1' missing after save.");

            Paragraph finalParagraph = FindContainingParagraph(finalBookmark.BookmarkStart);
            OfficeMath finalMath = FindTopLevelOfficeMath(finalParagraph);
            if (finalMath == null)
                throw new InvalidOperationException("No OfficeMath found in the paragraph of 'eq1' after replacement.");
        }
    }

    // Inserts a simple equation using a deterministic EQ‑field bootstrap workflow.
    // Uses a safe EQ switch (fraction) that reliably converts to OfficeMath.
    private static void InsertEquation(DocumentBuilder builder)
    {
        // Insert an EQ field.
        FieldEQ field = (FieldEQ)builder.InsertField(FieldType.FieldEquation, true);

        // Write a safe EQ argument into the field separator.
        builder.MoveTo(field.Separator);
        builder.Write(@"\f(1,2)");

        // Update the field so that the result is calculated before conversion.
        field.Update();

        // Return the builder to the field start location.
        builder.MoveTo(field.Start.ParentNode);

        // Convert the field to a real OfficeMath object.
        OfficeMath officeMath = field.AsOfficeMath();

        // Fallback: if conversion failed (unlikely with the safe pattern), try a simpler argument.
        if (officeMath == null)
        {
            builder.MoveTo(field.Separator);
            builder.Write(@"x+y");
            field.Update();
            builder.MoveTo(field.Start.ParentNode);
            officeMath = field.AsOfficeMath();
        }

        if (officeMath == null)
            throw new InvalidOperationException("EQ conversion failed for safe argument.");

        // Replace the field with the OfficeMath node.
        field.Start.ParentNode.InsertBefore(officeMath, field.Start);
        field.Remove();
    }

    // Walks up the node hierarchy to find the containing Paragraph.
    private static Paragraph FindContainingParagraph(Node node)
    {
        while (node != null && !(node is Paragraph))
            node = node.ParentNode;
        return node as Paragraph;
    }

    // Retrieves the top‑level OfficeMath (MathObjectType.OMathPara) within a paragraph.
    private static OfficeMath FindTopLevelOfficeMath(Paragraph paragraph)
    {
        if (paragraph == null) return null;
        return paragraph.GetChildNodes(NodeType.OfficeMath, false)
                        .Cast<OfficeMath>()
                        .FirstOrDefault(m => m.MathObjectType == MathObjectType.OMathPara);
    }
}
