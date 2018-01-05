using System;
using System.Collections.Generic;
using System.IO;

namespace docx2txt
{
    class Program
    {
        static DocxConverter _converter = new DocxConverter();

        static void Main(string[] args)
        {
            if (args.Length != 2)
            {
                PrintUsage();
                return;
            }
            
            switch (args[0])
            {
                case "-f":
                    ConvertDocument(args[1]);
                    return;
                case "-d":
                    BatchConvert(Directory.EnumerateFiles(args[1], "*.docx"));
                    return;
                default:
                    PrintUsage();
                    return;
            }
        }

        static bool ConvertDocument(string input)
        {
            Console.WriteLine($"Converting {input} ...");
            try
            {
                _converter.Convert(input, input.Replace(".docx", ".txt"));
                Console.WriteLine("Done");
                return true;
            }
            catch (Exception e)
            {
                Console.WriteLine($"Error converting {input}: {e.ToString()}");
                return false;
            }
        }

        static void BatchConvert(IEnumerable<string> inputs)
        {
            var errors = new List<string>();
            foreach (var input in inputs)
            {
                if (!ConvertDocument(input))
                {
                    errors.Add(input);
                }
            }

            Console.WriteLine("Documents converted");
            if (errors.Count > 0)
            {
                Console.WriteLine("It didn't work for the following documents:");
                Console.WriteLine(String.Join(Environment.NewLine, errors));
            }
        }

        static void PrintUsage()
        {
            Console.WriteLine("docx2txt converts a Word document to a UTF8 txt file.");
            Console.WriteLine();
            Console.WriteLine("Usage:");
            Console.WriteLine("    docx2txt -f <input file>          Converts a single file");
            Console.WriteLine("    docx2txt -d <input dir>           Converts all documents in directory");
        }
    }
}
