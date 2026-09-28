using System;
using System.IO;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Fields;
using Aspose.Words.Math;
using Aspose.Words.Tables;

public class OfficeMathReplaceExample
{
    public static void Main()
    {
        const string inputPath = "sample.docx";
        const string outputPath = "modified.docx";

        // 1. Create a sample DOCX with two identifiable equations.
        CreateSampleDocument(inputPath);

        // 2. Load the document.
        Document doc = new Document(inputPath);

        // 3. Replace the equation identified by bookmark "Eq1".
        ReplaceEquationByBookmark(doc, "Eq1", @"\r(5,x)");

        // 4. Save the modified document.
        doc.Save(outputPath, SaveFormat.Docx);

        // 5. Validate that the output file exists.
        if (!File.Exists(outputPath))
            throw new InvalidOperationException($"Failed to create output file: {outputPath}");

        // 6. Simple validation that the new equation exists.
        Document verifyDoc = new Document(outputPath);
        Bookmark bm = verifyDoc.Range.Bookmarks["Eq1"];
        Paragraph para = bm?.BookmarkStart?.ParentNode as Paragraph;
        OfficeMath math = para?.GetChildNodes(NodeType.OfficeMath, false)
                               .Cast<OfficeMath>()
                               .FirstOrDefault();
        if (math == null)
            throw new InvalidOperationException("Replacement equation was not found.");

        Console.WriteLine("Equation replacement succeeded.");
    }

    private static void CreateSampleDocument(string path)
    {
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // First equation with bookmark "Eq1"
        builder.Writeln("Paragraph before first equation.");
        builder.StartBookmark("Eq1");
        InsertEquationViaEQ(builder, @"\f(1,2)"); // Simple fraction 1/2
        builder.EndBookmark("Eq1");
        builder.Writeln();

        // Second equation with bookmark "Eq2"
        builder.Writeln("Paragraph before second equation.");
        builder.StartBookmark("Eq2");
        InsertEquationViaEQ(builder, @"\r(3,x)"); // Simple root
        builder.EndBookmark("Eq2");
        builder.Writeln();

        doc.Save(path, SaveFormat.Docx);
    }

    private static void InsertEquationViaEQ(DocumentBuilder builder, string eqSwitch)
    {
        // Insert an EQ field.
        Field field = builder.InsertField(FieldType.FieldEquation, true);
        // Write the EQ argument.
        builder.MoveTo(field.Separator);
        builder.Write(eqSwitch);
        // Update the field so that it can be converted to OfficeMath.
        field.Update();

        // Convert to OfficeMath.
        FieldEQ fieldEq = field as FieldEQ;
        if (fieldEq == null)
            throw new InvalidOperationException("Inserted field is not a FieldEQ.");

        OfficeMath officeMath = fieldEq.AsOfficeMath();
        if (officeMath == null)
            throw new InvalidOperationException("EQ field could not be converted to OfficeMath.");

        // Insert the OfficeMath node before the field start.
        Node startNode = field.Start;
        startNode.ParentNode.InsertBefore(officeMath, startNode);
        // Remove the original field.
        field.Remove();
    }

    private static void ReplaceEquationByBookmark(Document doc, string bookmarkName, string newEqSwitch)
    {
        // Locate the bookmark.
        Bookmark bookmark = doc.Range.Bookmarks[bookmarkName];
        if (bookmark == null)
            throw new InvalidOperationException($"Bookmark '{bookmarkName}' not found.");

        // The bookmark should enclose a paragraph that contains the OfficeMath node.
        Paragraph para = bookmark.BookmarkStart?.ParentNode as Paragraph;
        if (para == null)
            throw new InvalidOperationException($"Bookmark '{bookmarkName}' is not inside a paragraph.");

        // Find the existing OfficeMath node within the paragraph.
        OfficeMath existingMath = para.GetChildNodes(NodeType.OfficeMath, false)
                                    .Cast<OfficeMath>()
                                    .FirstOrDefault();
        if (existingMath == null)
            throw new InvalidOperationException($"No OfficeMath found within bookmark '{bookmarkName}'.");

        // Move the builder to the existing OfficeMath node.
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.MoveTo(existingMath);

        // Insert the new equation using the deterministic EQ-field bootstrap.
        InsertEquationViaEQ(builder, newEqSwitch);

        // Remove the old equation.
        existingMath.Remove();
    }
}
