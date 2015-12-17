using Fclp;
using System;
using System.IO;
using System.Linq;
using System.Text;
using MarkdownSharp;
using System.Reflection;
using DevExpress.Export.Xl;
using System.Collections.Generic;

namespace SessionBuilder
{
    class Program
    {
        private static string _path;
        private static string _output;
        private static readonly List<Record> _records = new List<Record>();
        private static bool _make = false;
        static void Main(string[] args)
        {
            bool help = false;
            bool raw = false;
            var p = new FluentCommandLineParser();
            string brk = "\n\t\t";

            p.Setup<string>('p', "path")
             .Callback(pth => _path = pth)
             .WithDescription(string.Format("{0}Search path:{0} Raw mode,{0}  .../1_Monday/NameOfInterview/<filename>.md{0} Final mode,{0}  .../1_Monday/NameOfInterview.mp4{0}  (looks for corresponding NameofInterview.md)", brk))
             .Required();

            p.Setup<bool>('r', "raw")
             .Callback(r => raw = r)
             .WithDescription(brk + "Final encoded files or raw? (default is final)")
             .SetDefault(false);

            p.Setup<bool>('m', "make")
             .Callback(r => _make = r)
             .WithDescription(brk + "Make default md files (if they don't exist)")
             .SetDefault(false);

            p.Setup<string>('o', "output")
             .Callback(o => _output = o)
             .WithDescription(brk + "Excel output location")
             .SetDefault(string.Empty);

            p.SetupHelp("?", "help")
             .Callback(text =>
             {
                 Console.WriteLine("\nThis program recursively looks for videos and a corresponding\nmarkdown file to create a title and description excel spreadsheet.\nThe first line of the markdown file is expected to contain the\nvideo title. The rest is considered the description.");
                 Console.WriteLine(text);
                 Console.WriteLine("\nExample:{0}SessionBuilder.exe -p C:\\folder -o C:\\folder\\Sessions.xlsx -f", brk);
                 help = true;
             });



            p.Parse(args);

            if (!help)
            {
                if (!Directory.Exists(_path))
                {
                    WriteLine("Valid path is required...", ConsoleColor.Red);
                    p.HelpOption.ShowHelp(p.Options);
                    return;
                }


                if (_output == string.Empty)
                    _output = Path.Combine(_path, "sessions.xlsx");


                Console.WriteLine("\nRecursively reading in {0} mode...", raw ? "raw" : "final");
                Console.WriteLine(_path);
                _records.Clear();
                Crawl(_path, raw);
                if (_records.Count > 0)
                {
                    WriteLine(string.Format("\nWriting spreadsheet {0}...", _output), ConsoleColor.Green);
                    CreateSpreadsheet(_records, _output);
                }
                else
                    WriteLine("\nNo descriptions found...", ConsoleColor.Red);
                Console.WriteLine("\nDone");
            }
        }

        private static void Crawl(string folder, bool raw, string pad = "\t")
        {
            foreach (var item in Directory.EnumerateDirectories(folder))
            {
                var breakup = item.Split('\\');
                Console.WriteLine(pad + breakup[breakup.Length - 1]);
                Crawl(item, raw, pad + "\t");
            }

            var pattern = raw ? "*.md" : "*.mp4";
            foreach (var item in Directory.EnumerateFiles(folder, pattern))
            {
                // FINAL
                // .../1_Monday/NameOfInterview.mp4
                // RAW
                // .../1_Monday/NameOfInterview/<filename>.md

                var path = Path.GetDirectoryName(item).Split('\\');
                Record r = new Record();
                string md = string.Empty;
                string parent = "";
                if (raw)
                {
                    md = item;
                    parent = path[path.Length - 2];
                    r.FileName = string.Format("{0}.mp4", path[path.Length - 1]);
                }
                else
                {
                    md = item.Replace(".mp4", ".md");
                    WriteLine(pad + Path.GetFileName(item), ConsoleColor.Cyan);
                    // no description, no point
                    if (!File.Exists(md))
                    {
                        if (_make)
                        {
                            WriteLine(string.Format("{0}Writing out default markdown to {1}", pad, Path.GetFileName(md)), ConsoleColor.Green);
                            File.WriteAllText(md, "{ tag1, tag2, tag3 }\n# Title Here\n\nDescription here.");
                        }
                        continue;
                    }
                    parent = path[path.Length - 1];
                    r.FileName = Path.GetFileName(item);
                    r.FilePath = item;
                    r.RelativePath = item.Replace(_path + "\\", "");
                }

                var data = parent.Split('_');
                r.Session = data[0];
                r.Day = data[1];

                r.LoadFile(md);

                _records.Add(r);

                if(r.Title == "Title Here")
                    WriteLine(string.Format("{0}Found: \"{1}\"", pad + (raw ? "" : "\t"), r.Title), ConsoleColor.Red);
                else
                    WriteLine(string.Format("{0}Found: \"{1}\"", pad + (raw ? "" : "\t"), r.Title), ConsoleColor.Green);
            }
        }

        private static void CreateSpreadsheet(IEnumerable<Record> records, string file)
        {
            if (File.Exists(file))
                File.Delete(file);
            // Create an exporter instance. 
            IXlExporter exporter = XlExport.CreateExporter(XlDocumentFormat.Xlsx);

            // Create the FileStream object with the specified file path. 
            using (FileStream stream = new FileStream(file, FileMode.Create, FileAccess.ReadWrite))
            {
                // Create a new document and begin to write it to the specified stream. 
                using (IXlDocument document = exporter.CreateDocument(stream))
                {
                    var props = typeof(Record).GetProperties();
                    // Add a new worksheet to the document.
                    using (IXlSheet sheet = document.CreateSheet())
                    {

                        // Specify the worksheet name.
                        sheet.Name = "Sessions";

                        XlCellFormatting cellFormatting = new XlCellFormatting();
                        cellFormatting.Font = new XlFont();
                        cellFormatting.Font.Bold = true;

                        // Create the header row.
                        using (IXlRow row = sheet.CreateRow())
                        {
                            // create cell
                            foreach (var prop in props)
                            {

                                using (IXlCell cell = row.CreateCell())
                                {
                                    cell.Value = prop.Name;
                                    cell.ApplyFormatting(cellFormatting);
                                }
                            }
                        }

                        foreach (var record in records)
                        {
                            using (IXlRow row = sheet.CreateRow())
                            {
                                // create cell
                                foreach (var prop in props)
                                {
                                    using (IXlCell cell = row.CreateCell())
                                        cell.Value = prop.GetValue(record)?.ToString();
                                }
                            }
                        }
                    }
                }

            }
        }

        private static string GetCurrentPath()
        {
            string codeBase = Assembly.GetExecutingAssembly().CodeBase;
            UriBuilder uri = new UriBuilder(codeBase);
            var path = Uri.UnescapeDataString(uri.Path);
            return Path.GetDirectoryName(path);
        }

        private static void WriteLine(string text, ConsoleColor color)
        {
            var curColor = Console.ForegroundColor;
            Console.ForegroundColor = color;
            Console.WriteLine(text);
            Console.ForegroundColor = curColor;
        }

        private static void Write(string text, ConsoleColor color)
        {
            var curColor = Console.ForegroundColor;
            Console.ForegroundColor = color;
            Console.Write(text);
            Console.ForegroundColor = curColor;
        }
    }

    public class Record
    {
        public string Title { get; set; }

        public string Description { get; set; }

        public string Tags { get; set; }

        public string MarkdownDescription { get; set; }

        public string Day { get; set; }

        public string Session { get; set; }

        public string FileName { get; set; }

        public string FilePath { get; set; }

        public string RelativePath { get; set; }

        public override string ToString()
        {
            StringBuilder sb = new StringBuilder();
            foreach (var p in GetType().GetProperties())
                sb.AppendLine(string.Format("{0}: {1}", p.Name, p.GetValue(this)));

            return sb.ToString();
        }

        public void LoadFile(string file)
        {
            if (File.Exists(file))
            {
                using (var f = File.OpenText(file))
                {
                    var text = f.ReadLine();
                    if (text.StartsWith("{"))
                    {
                        Tags = text.Replace("{", "").Replace("}", "").Trim();
                        text = f.ReadLine();
                    }

                    Title = text.Replace("#", "").Trim();

                    MarkdownDescription = f.ReadToEnd();
                    Description = new Markdown()
                                    .Transform(MarkdownDescription);
                }
            }
        }
    }


}
