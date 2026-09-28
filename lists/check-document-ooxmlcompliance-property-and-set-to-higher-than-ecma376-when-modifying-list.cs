using System;
using System.IO;

namespace OoxmlDemo
{
    // Minimal enum to represent compliance levels
    public enum OoxmlCompliance
    {
        Ecma376,
        Office2010,
        Office2013,
        Office2016
    }

    // Stub for a Wordprocessing document
    public sealed class WordprocessingDocument : IDisposable
    {
        public OoxmlCompliance OoxmlCompliance { get; set; }

        private readonly string _filePath;
        private bool _disposed;

        private WordprocessingDocument(string filePath)
        {
            _filePath = filePath;
            OoxmlCompliance = OoxmlCompliance.Ecma376; // default value
        }

        // Factory method mimicking the OpenXML SDK API
        public static WordprocessingDocument Create(string filePath, WordprocessingDocumentType type)
        {
            // In a real scenario the file would be created here.
            // For this demo we just return a new instance.
            Console.WriteLine($"Creating WordprocessingDocument at '{filePath}' with type {type}.");
            return new WordprocessingDocument(filePath);
        }

        // Adds the main document part and returns a stub object
        public MainDocumentPart AddMainDocumentPart()
        {
            Console.WriteLine("Adding MainDocumentPart.");
            return new MainDocumentPart();
        }

        public void Dispose()
        {
            if (!_disposed)
            {
                // In a real scenario the package would be closed here.
                Console.WriteLine($"Disposing WordprocessingDocument for '{_filePath}'.");
                _disposed = true;
            }
        }
    }

    // Stub enum to match the SDK signature
    public enum WordprocessingDocumentType
    {
        Document,
        Template,
        MacroEnabledDocument,
        MacroEnabledTemplate
    }

    // Stub for the main document part
    public class MainDocumentPart
    {
        public Document Document { get; set; }

        // Adds a new part of the requested type
        public T AddNewPart<T>() where T : new()
        {
            Console.WriteLine($"Adding new part of type {typeof(T).Name}.");
            return new T();
        }

        // Simulates saving the document
        public void Save()
        {
            Console.WriteLine("Saving MainDocumentPart.");
        }
    }

    // Stub for the document body
    public class Document
    {
        public Body Body { get; }

        public Document(Body body)
        {
            Body = body;
        }
    }

    // Empty body placeholder
    public class Body { }

    // Stub for numbering definitions part
    public class NumberingDefinitionsPart
    {
        public Numbering Numbering { get; set; }
    }

    // Minimal representation of numbering structures
    public class Numbering
    {
        public AbstractNum AbstractNum { get; }
        public NumberingInstance NumberingInstance { get; }

        public Numbering(AbstractNum abstractNum, NumberingInstance numberingInstance)
        {
            AbstractNum = abstractNum;
            NumberingInstance = numberingInstance;
        }
    }

    public class AbstractNum
    {
        public Level Level { get; }
        public int AbstractNumberId { get; set; }

        public AbstractNum(Level level)
        {
            Level = level;
        }
    }

    public class Level
    {
        public NumberingFormat NumberingFormat { get; }
        public LevelText LevelText { get; }
        public StartNumberingValue StartNumberingValue { get; }

        public Level(NumberingFormat format, LevelText text, StartNumberingValue start)
        {
            NumberingFormat = format;
            LevelText = text;
            StartNumberingValue = start;
        }
    }

    public class NumberingFormat
    {
        public NumberFormatValues Val { get; set; }
    }

    public enum NumberFormatValues
    {
        Decimal,
        UpperRoman,
        LowerLetter
    }

    public class LevelText
    {
        public string Val { get; set; }
    }

    public class StartNumberingValue
    {
        public int Val { get; set; }
    }

    public class NumberingInstance
    {
        public AbstractNumId AbstractNumId { get; }
        public int NumberID { get; set; }

        public NumberingInstance(AbstractNumId abstractNumId)
        {
            AbstractNumId = abstractNumId;
        }
    }

    public class AbstractNumId
    {
        public int Val { get; set; }
    }

    public class Program
    {
        public static void Main()
        {
            string filePath = "example.docx";

            // Create a new Wordprocessing document
            using (WordprocessingDocument wordDoc = WordprocessingDocument.Create(filePath, WordprocessingDocumentType.Document))
            {
                // Add the main document part
                MainDocumentPart mainPart = wordDoc.AddMainDocumentPart();
                mainPart.Document = new Document(new Body());

                // Check the OoxmlCompliance property
                Console.WriteLine($"Initial OoxmlCompliance: {wordDoc.OoxmlCompliance}");
                if (wordDoc.OoxmlCompliance == OoxmlCompliance.Ecma376)
                {
                    // Set to a higher compliance level (Office2010)
                    wordDoc.OoxmlCompliance = OoxmlCompliance.Office2010;
                    Console.WriteLine($"OoxmlCompliance upgraded to: {wordDoc.OoxmlCompliance}");
                }

                // Add a NumberingDefinitionsPart (list definitions)
                NumberingDefinitionsPart numberingPart = mainPart.AddNewPart<NumberingDefinitionsPart>();
                numberingPart.Numbering = new Numbering(
                    new AbstractNum(
                        new Level(
                            new NumberingFormat { Val = NumberFormatValues.Decimal },
                            new LevelText { Val = "%1." },
                            new StartNumberingValue { Val = 1 })
                    )
                    {
                        AbstractNumberId = 1
                    },
                    new NumberingInstance(
                        new AbstractNumId { Val = 1 })
                    {
                        NumberID = 1
                    });

                // Save the main document part
                mainPart.Save();
            }

            // Clean up the created file (no real file was written in this stub)
            if (File.Exists(filePath))
            {
                File.Delete(filePath);
                Console.WriteLine($"Deleted temporary file '{filePath}'.");
            }
            else
            {
                Console.WriteLine($"No physical file created; cleanup not required.");
            }
        }
    }
}
