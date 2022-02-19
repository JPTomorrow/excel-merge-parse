using System;
using System.Data;
using System.IO;
using System.Linq;
using System.Reflection;
using JPMorrow.ExcelMerge;

namespace ExcelMergeParse
{
    class Program
    {
        public static bool header_skip = false;
        public static string[] cmds = new[] { "merge", "search", "clear" };
        public static string app_path { get; } = Directory.GetCurrentDirectory();
        public static string data_folder_path { get; } = app_path + "\\excel_parse_merge_data";
        public static string data_file_name { get; } = "current_data.json";

        public static MergeData current_data = null;

        [STAThread]
        static void Main(string[] args)
        {
            if(!Directory.Exists(data_folder_path))
                Directory.CreateDirectory(data_folder_path);

            if(File.Exists(data_folder_path + "\\" + data_file_name)) {
                current_data = MergeData.LoadData(data_folder_path + "\\" + data_file_name);
                Console.WriteLine("Excel data has been loaded and is ready to search through.\n");
            }

            bool quit = false;
            while(!quit) {
                quit = Update();
            }
        }

        public static bool Update() {
            if(header_skip == false) {
                PrintHeader();
                header_skip = true;
            }

            var cmd = PostCmd("Please enter a command :>> ").Trim();

            if(cmd.Equals("quit"))
                return true;
            else if(cmd.Equals("-h"))
                PrintHelp();
            else if(cmd.Equals("clear")) {
                Console.Clear();
            }
            else if(cmd.Equals("merge")) {
                MergeData data = ExMerge.MergeExcelFilesIntoTable();
                data.SaveData(data_folder_path + "\\" + data_file_name);

                Console.WriteLine(data.Table.Count() + " rows found.\n");
                Console.WriteLine("Data has been merged and is ready for search.\n");

                var rows_samples = data.Table.Take(data.Table.Count < 10 ? data.Table.Count : 10);

                foreach(var row in rows_samples) {
                    Console.WriteLine(string.Join(" █ ", row));
                }

                current_data = data;
            }
            else if(cmd.Contains("search")) {
                if(current_data == null)
                    Console.WriteLine("You must merge some excel files first before you can search them.");

                var terms = cmd.Split(' ').Where(x => !x.Equals("search"));
                var rows = current_data.FindRows(terms.ToArray());

                Console.WriteLine("Found " + rows.Count().ToString() + " rows.\n");

                foreach(var row in rows) {
                    Console.WriteLine(string.Join(" █ ", row));
                }
            }
            else {
                Console.WriteLine(cmd + " is not a valid command, please try again, or type -h for help.");
            }

            return false;
        }

        public static string PostCmd(string question) {
            Console.Write(question);
            var read = Console.ReadLine().ToLower();
            Console.WriteLine("\n");
            return read;
        }

        public static void PrintHeader() {
            Console.WriteLine("Welcome To Excel Merge Parser!\n");
            Console.WriteLine("This program is designed to let you select mutiple excel files,\nmerge them together, and then search for entries\n");
            Console.WriteLine("Type -h for help menu.\n");
        }

        public static void PrintHelp() {
            Console.WriteLine("merge | 'merge two or more excel files and save the output as json that will be read in by the program every time you boot.'");
            Console.WriteLine("search | 'search for row(s) by their text content.'");
            Console.WriteLine("clear | 'clear the console screen.'\n");
        }
    }
}
