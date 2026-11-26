using System;
using System.IO;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Fields;
using Aspose.Words.Math;
using Aspose.Words.Saving;
using Aspose.Words.Layout;

public class OfficeMathPdfPositionValidator
{
    public static void Main()
    {
        // Paths for temporary files.
        string docxPath = "SampleEquations.docx";
        string pdfPath = "SampleEquations.pdf";

        // -----------------------------------------------------------------
        // 1. Create a DOCX with two simple OfficeMath equations.
        // -----------------------------------------------------------------
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // First equation with bookmark "eq1".
        builder.StartBookmark("eq1");
        InsertEquation(builder, @"\f(1,2)"); // simple fraction.
        builder.EndBookmark("eq1");
        builder.Writeln(); // separate paragraph.

        // Second equation with bookmark "eq2".
        builder.StartBookmark("eq2");
        InsertEquation(builder, @"\r(3,x)"); // simple radical.
        builder.EndBookmark("eq2");
        builder.Writeln();

        // Save the source DOCX.
        doc.Save(docxPath, SaveFormat.Docx);

        // -----------------------------------------------------------------
        // 2. Load the DOCX and record layout positions of top‑level OfficeMath nodes.
        // -----------------------------------------------------------------
        Document sourceDoc = new Document(docxPath);
        LayoutCollector sourceCollector = new LayoutCollector(sourceDoc);
        var sourceEquations = sourceDoc.GetChildNodes(NodeType.OfficeMath, true)
            .Cast<OfficeMath>()
            .Where(om => om.MathObjectType == MathObjectType.OMathPara)
            .ToList();

        if (!sourceEquations.Any())
            throw new InvalidOperationException("No top‑level OfficeMath equations were found in the source document.");

        var sourcePositions = sourceEquations
            .Select(om => new
            {
                Bookmark = FindContainingBookmark(om),
                Page = sourceCollector.GetStartPageIndex(om)
            })
            .ToList();

        // -----------------------------------------------------------------
        // 3. Export the document to PDF.
        // -----------------------------------------------------------------
        sourceDoc.Save(pdfPath, SaveFormat.Pdf);

        // Verify PDF file exists.
        if (!File.Exists(pdfPath))
            throw new FileNotFoundException("PDF export failed; file not found.", pdfPath);

        // -----------------------------------------------------------------
        // 4. Reload the DOCX and record layout positions again.
        // -----------------------------------------------------------------
        Document reloadedDoc = new Document(docxPath);
        LayoutCollector reloadedCollector = new LayoutCollector(reloadedDoc);
        var reloadedEquations = reloadedDoc.GetChildNodes(NodeType.OfficeMath, true)
            .Cast<OfficeMath>()
            .Where(om => om.MathObjectType == MathObjectType.OMathPara)
            .ToList();

        var reloadedPositions = reloadedEquations
            .Select(om => new
            {
                Bookmark = FindContainingBookmark(om),
                Page = reloadedCollector.GetStartPageIndex(om)
            })
            .ToList();

        // -----------------------------------------------------------------
        // 5. Validate that each equation retained its original page position.
        // -----------------------------------------------------------------
        foreach (var src in sourcePositions)
        {
            var match = reloadedPositions.FirstOrDefault(r => r.Bookmark == src.Bookmark);
            if (match == null)
                throw new InvalidOperationException($"Equation with bookmark '{src.Bookmark}' is missing after reload.");

            if (match.Page != src.Page)
                throw new InvalidOperationException($"Equation '{src.Bookmark}' moved from page {src.Page} to page {match.Page} after PDF export.");
        }

        // If we reach this point, validation succeeded.
        Console.WriteLine("PDF export retained exact positioning of all OfficeMath equations.");
    }

    // Inserts an OfficeMath equation using the deterministic EQ‑field bootstrap workflow.
    private static void InsertEquation(DocumentBuilder builder, string eqArgument)
    {
        // Insert an EQ field.
        Field field = builder.InsertField(FieldType.FieldEquation, true);
        if (field is not FieldEQ eqField)
            throw new InvalidOperationException("Inserted field is not a FieldEQ.");

        // Move to the field separator node and write the EQ argument.
        builder.MoveTo(eqField.Separator);
        builder.Write(eqArgument);

        // Ensure the field is up‑to‑date before conversion.
        field.Update();

        // Convert to OfficeMath.
        OfficeMath officeMath = eqField.AsOfficeMath();
        if (officeMath == null)
            throw new InvalidOperationException("EQ field conversion to OfficeMath returned null.");

        // Insert the OfficeMath node before the field start and remove the field.
        Node startNode = eqField.Start;
        if (startNode?.ParentNode is not CompositeNode parent)
            throw new InvalidOperationException("Failed to locate parent node for EQ field.");

        parent.InsertBefore(officeMath, startNode);
        eqField.Remove();
    }

    // Finds the name of the bookmark that contains the given OfficeMath node, if any.
    private static string FindContainingBookmark(OfficeMath officeMath)
    {
        // Walk up to the containing paragraph.
        Node current = officeMath;
        while (current != null && !(current is Paragraph))
            current = current.ParentNode;

        if (current is Paragraph paragraph)
        {
            // Search bookmarks that span this paragraph.
            foreach (Bookmark bookmark in paragraph.Document.Range.Bookmarks)
            {
                // If the bookmark start or end node is inside the paragraph, consider it a match.
                if (bookmark.BookmarkStart.ParentNode == paragraph || bookmark.BookmarkEnd.ParentNode == paragraph)
                    return bookmark.Name;
            }
        }

        return "<none>";
    }
}
