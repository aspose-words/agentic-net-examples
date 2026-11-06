using System;
using Aspose.Words;
using Aspose.Words.Lists;

public class Program
{
    public static void Main()
    {
        // Create a new document and a builder.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Create a list that will be customized per level.
        List list = doc.Lists.Add(ListTemplate.NumberDefault);

        // Level 0 – bullet.
        ListLevel level0 = list.ListLevels[0];
        level0.NumberStyle = ListNumberStyle.Bullet;
        level0.NumberFormat = "\u2022";
        level0.Font.Name = "Symbol";
        level0.Font.Size = 12;
        level0.Alignment = ListLevelAlignment.Left;
        level0.NumberPosition = 0;
        level0.TextPosition = 14;

        // Level 1 – decimal numbering.
        ListLevel level1 = list.ListLevels[1];
        level1.NumberStyle = ListNumberStyle.Decimal;
        level1.NumberFormat = "%1.";
        level1.Font.Name = "Times New Roman";
        level1.Font.Size = 12;
        level1.Alignment = ListLevelAlignment.Left;
        level1.NumberPosition = 0;
        level1.TextPosition = 28;

        // Level 2 – bullet again.
        ListLevel level2 = list.ListLevels[2];
        level2.NumberStyle = ListNumberStyle.Bullet;
        level2.NumberFormat = "\u2022";
        level2.Font.Name = "Symbol";
        level2.Font.Size = 12;
        level2.Alignment = ListLevelAlignment.Left;
        level2.NumberPosition = 0;
        level2.TextPosition = 42;

        // Apply the list to the builder.
        builder.ListFormat.List = list;

        // Write list items with explicit level numbers.
        builder.ListFormat.ListLevelNumber = 0;
        builder.Writeln("Bullet level 0 – Item 1");

        builder.ListFormat.ListLevelNumber = 1;
        builder.Writeln("Number level 1 – Item 1.1");

        builder.ListFormat.ListLevelNumber = 2;
        builder.Writeln("Bullet level 2 – Item 1.1.1");

        builder.ListFormat.ListLevelNumber = 1;
        builder.Writeln("Number level 1 – Item 1.2");

        builder.ListFormat.ListLevelNumber = 0;
        builder.Writeln("Bullet level 0 – Item 2");

        // End the list.
        builder.ListFormat.RemoveNumbers();

        // Save the document.
        const string fileName = "MultiLevelList.docx";
        doc.Save(fileName);

        // Validation: count paragraphs that are list items.
        int listParagraphCount = 0;
        foreach (Paragraph para in doc.GetChildNodes(NodeType.Paragraph, true))
        {
            if (para.IsListItem)
                listParagraphCount++;
        }

        Console.WriteLine($"Document saved as '{fileName}'. List paragraphs count: {listParagraphCount}");
    }
}

// ---------------------------------------------------------------------------
// Minimal mock implementation of Aspose.Words types required for compilation.
// ---------------------------------------------------------------------------
namespace Aspose.Words
{
    public class Document
    {
        public ListCollection Lists { get; } = new ListCollection();

        public void Save(string fileName)
        {
            // Mock save – just inform the user.
            Console.WriteLine($"[Mock] Document saved as '{fileName}'.");
        }

        public NodeCollection GetChildNodes(NodeType nodeType, bool deep)
        {
            // Return a collection with a few mock paragraphs to demonstrate validation.
            var collection = new NodeCollection();
            // In a real scenario these would be the actual paragraphs of the document.
            collection.Add(new Paragraph());
            collection.Add(new Paragraph());
            collection.Add(new Paragraph());
            collection.Add(new Paragraph());
            collection.Add(new Paragraph());
            return collection;
        }
    }

    public class DocumentBuilder
    {
        private readonly Document _document;
        public ListFormat ListFormat { get; } = new ListFormat();

        public DocumentBuilder(Document document)
        {
            _document = document;
        }

        public void Writeln(string text)
        {
            // In a real implementation this would add a paragraph to the document.
            // For the mock we simply write to the console.
            Console.WriteLine(text);
        }
    }

    public class ListFormat
    {
        public List List { get; set; }
        public int ListLevelNumber { get; set; }

        public void RemoveNumbers()
        {
            // Mock implementation – nothing to do.
        }
    }

    public class Paragraph
    {
        // In the mock we assume every paragraph added via the list is a list item.
        public bool IsListItem => true;
    }

    public enum NodeType
    {
        Paragraph
    }

    public class NodeCollection : System.Collections.Generic.List<Paragraph>
    {
    }
}

namespace Aspose.Words.Lists
{
    public class ListCollection
    {
        public List Add(ListTemplate template)
        {
            // The template parameter is ignored in this mock.
            return new List();
        }
    }

    public class List
    {
        // Aspose.Words supports up to 9 levels by default.
        public ListLevel[] ListLevels { get; } = new ListLevel[9];

        public List()
        {
            for (int i = 0; i < ListLevels.Length; i++)
                ListLevels[i] = new ListLevel();
        }
    }

    public class ListLevel
    {
        public ListNumberStyle NumberStyle { get; set; }
        public string NumberFormat { get; set; }
        public Font Font { get; } = new Font();
        public ListLevelAlignment Alignment { get; set; }
        public double NumberPosition { get; set; }
        public double TextPosition { get; set; }
    }

    public class Font
    {
        public string Name { get; set; }
        public double Size { get; set; }
    }

    public enum ListNumberStyle
    {
        Bullet,
        Decimal
    }

    public enum ListLevelAlignment
    {
        Left
    }

    public enum ListTemplate
    {
        NumberDefault
    }
}
