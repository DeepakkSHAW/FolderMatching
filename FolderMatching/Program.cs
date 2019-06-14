using CommandLine;
using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;

namespace FolderMatching
{
    //Declare some options
    public class Options
    {
        //Format:
        //[Option(char shortoption, string longoption, Required = bool,  MetaValue = string, Min = int, Seperator = char, SetName = string)]
        //public <type> <name> { get; set; }

        [Option('r', "read", Required = true, HelpText = "Filename", Separator = ' ')]
        public IEnumerable<string> InputFiles { get; set; }

        [Option('e', "echo", Required = true, HelpText = "Echo message")]
        public string echo { get; set; }

    }
    class Program
    {

        static void Main(string[] args)
        {
            var excelFilePath = @"c:\Temp\Pluralsight-Course-List.xlsx";
            var folderPath = @"D:\Pluralsight";

            Action<string> cw = Console.WriteLine;
            cw("****************Welcome to Folder matching****************");
            //if(args.Length > 0)
            if (args.Length > 0)
                excelFilePath = args[0] ?? excelFilePath;

            if (args.Length > 1)
                folderPath = args[1] ?? folderPath;

            var coursesInDisk = new List<string>();

            try {
                if (Directory.Exists(folderPath))
                    coursesInDisk = Directory.GetDirectories(folderPath).ToList();

                var excelCourse = ReadCourcesFromExcel(cw, excelFilePath);

                //** Phase1: Looking into Excel file if course doesn't exist **//
                foreach (string s in coursesInDisk)
                {
                    var courseInDisk = s.Remove(0, s.LastIndexOf(Path.DirectorySeparatorChar) + 1);
                    var found = excelCourse.FirstOrDefault(i => i.Contains(courseInDisk));

                    if (found == null)
                        cw($"The course '{courseInDisk}' missing into excel file.");

                }

                //** Phase2: Looking into File System if course doesn't exist **//
                foreach (string s in excelCourse)
                {
                    var found = coursesInDisk.FirstOrDefault(i => i.Contains(s));

                    if (found == null)
                        cw($"The course '{s}' missing in disk drive file.");

                }

            }
            catch (Exception ex)
            {
                cw(ex.Message);
            }


            cw("Press any key to close..");
            Console.ReadKey();
        }

        private static List<string> ReadCourcesFromExcel(Action<string> cw, string excelFilePath)
        {
            try
            {
                //Reading Course from Excel file
                var excelCourses = new List<string>();
                FileInfo file = new FileInfo(excelFilePath);
                using (ExcelPackage package = new ExcelPackage(file))
                {
                    //ExcelWorksheet worksheet = package.Workbook.Worksheets["Plurasite Course"];
                    //int rowCount = worksheet.Dimension.Rows;
                    //int ColCount = worksheet.Dimension.Columns;
                    var ws = package.Workbook.Worksheets.First();
                    var hasHeader = true;
                    var startRow = hasHeader ? 2 : 1;
                    for (int rowNum = startRow; rowNum <= ws.Dimension.End.Row; rowNum++)
                    {
                        var aCourse = ws.Cells[rowNum, 2].Value.ToString();
                        excelCourses.Add(aCourse);
                    //    cw(aCourse);
                    }
                }
                return excelCourses;
            }
            catch (Exception ex)
            {
                cw($"Error occurred: {ex.Message}");
                throw ex;
            }
        }
    }
}
