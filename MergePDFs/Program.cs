﻿using iTextSharp.text;
using iTextSharp.text.pdf;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;

namespace MergePDFs
{
    internal class Program
    {
        public static void AddPageNumber(IList<Dictionary<string, object>> dicts, Dictionary<string, string> namedPageDict, int increment)
        {
            for (var i = 0; i < dicts.Count; i++)
            {
                var Dict = dicts[i];
                if (Dict.ContainsKey("Page"))
                {
                    var Page = Dict["Page"] as string;
                    Dict["Page"] = (int.Parse(Page.Substring(0, Page.IndexOf(' '))) + increment) + Page.Substring(Page.IndexOf(' '));
                }
                else
                    Dict.Add("Page", "");

                if (Dict.ContainsKey("Named"))
                {
                    var Named = Dict["Named"] as string;
                    var Page = namedPageDict[Named].Replace("null", "0");

                    Dict["Page"] = (int.Parse(Page.Substring(0, Page.IndexOf(' '))) + increment) + Page.Substring(Page.IndexOf(' '));
                    Dict.Remove("Named");
                }

                if (Dict.ContainsKey("Kids"))
                    AddPageNumber(Dict["Kids"] as IList<Dictionary<string, object>>, namedPageDict, increment);
            }
        }

        public static void DeleteDirectory(string path)
        {
            foreach (string directory in Directory.GetDirectories(path))
            {
                DeleteDirectory(directory);
            }

            try
            {
                Directory.Delete(path, true);
            }
            catch (IOException)
            {
                Directory.Delete(path, true);
            }
            catch (UnauthorizedAccessException)
            {
                Directory.Delete(path, true);
            }
        }

        public static bool MergePDFs(IEnumerable<string> fileNames, string targetPdf)
        {
            bool merged = true;
            var Bookmark = new List<Dictionary<string, object>>();
            using (var stream = new MemoryStream())
            {
                var document = new Document();
                var pdf = new PdfCopy(document, stream);
                PdfReader reader = null;
                try
                {
                    document.Open();
                    var NumberOfPages = 0;
                    foreach (string file in fileNames)
                    {
                        var FileName = System.IO.Path.GetFileName(file);
                        using (reader = new PdfReader(file))
                        {
                            var IndividualBookmark = SimpleBookmark.GetBookmark(reader);
                            if (IndividualBookmark != null)
                            {
                                AddPageNumber(IndividualBookmark, SimpleNamedDestination.GetNamedDestination(reader, false), NumberOfPages);
                            }

                            Bookmark.Add(new Dictionary<string, object> {
                                   {"Title", FileName.Substring(0, FileName.IndexOf('.'))},
                                   {"Open", "false"},
                                   {"Page", (NumberOfPages + 1) + " Fit"},
                                   {"Action", "GoTo"},
                                   {"Kids", IndividualBookmark}
                                });

                            NumberOfPages += reader.NumberOfPages;
                            pdf.AddDocument(reader);
                        }
                    }
                }
                catch (Exception e)
                {
                    Console.WriteLine(e.Message);
                    merged = false;
                    if (reader != null)
                    {
                        reader.Close();
                    }

                    throw;
                }
                finally
                {
                    if (document != null)
                    {
                        document.Close();
                    }
                }

                using (var NewReader = new PdfReader(stream.ToArray()))
                using (var fileStream = new FileStream(targetPdf, FileMode.Create))
                using (var Stamper = new PdfStamper(NewReader, fileStream))
                {
                    stream.Dispose();
                    Stamper.Outlines = Bookmark;
                }
            }

            return merged;
        }

        public static string MoveFolder(string folder, string path)
        {
            var FolderName = System.IO.Path.GetFileName(folder);
            var FolderSpaceIndex = FolderName.LastIndexOf(' ');
            if (FolderSpaceIndex != -1)
            {
                var Rest = FolderName.Substring(Math.Min(FolderName.Length, FolderSpaceIndex + 1));
                int test;
                if (int.TryParse(Rest, out test))
                    FolderName = FolderName.Substring(0, FolderSpaceIndex);
            }

            var FolderNumbers = System.IO.Directory.EnumerateDirectories(path)
                .Select(x => System.IO.Path.GetFileName(x).ToLower())
            .Where(x => x.Contains(FolderName.ToLower()) && x.Replace(FolderName.ToLower(), "").AsEnumerable().All(y => Char.IsDigit(y) || y == ' '))
            .Select(x =>
            {
                var SpaceIndex = x.LastIndexOf(' ');
                if (SpaceIndex == -1)
                    return 0;
                int result;
                int.TryParse(x.Substring(SpaceIndex), out result);
                return result;
            })
            .ToArray();

            var NextNumber = Enumerable.Range(1, FolderNumbers.Length + 1).Except(FolderNumbers).Min();

            var Suffix = FolderNumbers.Any() ? " " + NextNumber.ToString().PadLeft(3, '0') : "";

            Directory.Move(folder, path + "\\" + FolderName + Suffix);
            return path + "\\" + FolderName + Suffix;
        }

        private static void Main(string[] args)
        {
            if (args.Length != 1)
            {
                Console.WriteLine(@"mergepdfs ""path folder containing the pdfs""");
                return;
            }

            var Folder = args[0].TrimEnd(new[] { '\\' });

            var PDFFiles = Directory.GetFiles(Folder, "*.pdf").OrderBy(x => Path.GetFileName(x));
            if (!PDFFiles.Any())
            {
                Console.WriteLine("There are no pdf files in " + Folder);
                return;
            }
            var FolderName = Path.GetFileName(Folder);
            var ParentPath = Directory.GetParent(Directory.GetParent(Folder).ToString()).ToString();

            if (!Directory.Exists(ParentPath + "\\Success"))
                Directory.CreateDirectory(ParentPath + "\\Success");

            try
            {
                var FolderNumbers = System.IO.Directory.EnumerateDirectories(ParentPath + "\\Success")
                    .Select(x => System.IO.Path.GetFileName(x).ToLower())
                .Where(x => x.Contains(FolderName.ToLower()) && x.Replace(FolderName.ToLower(), "").AsEnumerable().All(y => Char.IsDigit(y) || y == ' '))
                .Select(x =>
                {
                    var SpaceIndex = x.LastIndexOf(' ');
                    if (SpaceIndex == -1)
                        return 0;
                    int result;
                    int.TryParse(x.Substring(SpaceIndex), out result);
                    return result;
                })
                .ToArray();

                var Suffix = FolderNumbers.Any() ? " " + (FolderNumbers.Max() + 1).ToString().PadLeft(3, '0') : "";
                if (!Directory.Exists(ParentPath + "\\Success\\" + FolderName + Suffix))
                    Directory.CreateDirectory(ParentPath + "\\Success\\" + FolderName + Suffix);
                MergePDFs(PDFFiles, ParentPath + "\\Success\\" + FolderName + Suffix + "\\" + FolderName + ".pdf");
                DeleteDirectory(Folder);
            }
            catch (Exception e)
            {
                Console.WriteLine(e.Message);
                if (Directory.Exists(ParentPath + "\\Success\\" + FolderName) && !Directory.EnumerateFiles(ParentPath + "\\Success\\" + FolderName).Any())
                    DeleteDirectory(ParentPath + "\\Success\\" + FolderName);

                if (!Directory.Exists(ParentPath + "\\Errors"))
                    Directory.CreateDirectory(ParentPath + "\\Errors");

                if (!Directory.Exists(ParentPath + "\\Logs"))
                    Directory.CreateDirectory(ParentPath + "\\Logs");

                File.WriteAllText(ParentPath + "\\Logs\\" + Path.GetFileName(MoveFolder(Folder, ParentPath + "\\Errors")) + ".txt", e.ToString());
            }
        }
    }
}