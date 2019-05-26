using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Excel_Sharp;
using ClosedXML.Excel;
using System.IO;
using System.Text.RegularExpressions;



namespace Easy_XLS
{

    struct sb
    {
        public string name;
        public string sfz;
        public int zhengfan;
    }

    class Program
    {

        public static string shaixuan_name(string str)
        {
            Regex ConnoteA = new Regex(@"[\u4e00-\u9fa5].*?(?=-)");

            Match m = ConnoteA.Match(str);
            return m.Value;
        }
        public static string shaixuan_zhengfan(string str)
        {
            Regex ConnoteA = new Regex(@"(?<=-).*?(?=-)");

            Match m = ConnoteA.Match(str);

            return m.Value;
        }

        public static string shaixuan_sfz(string str)
        {
            Regex ConnoteA = new Regex(@"(^?\d{18}?$)|(^?\d{15}$?)");
            Match m = ConnoteA.Match(str);
            return m.Value;
        }
        static void Main(string[] args)
        {

            string help = "";
            help += "-i 身份证路径"+"\n";
            help += "-t 文本文档路径" + "\n";
            help += "-h 查看帮助" + "\n";
            help += "-o excel文档路径" + "\n";


            string sfz_path="";

            string excel_path="";
            
            string txt_path="";

#if DEBUG
            excel_path = @"..\..\1.xlsx";
            txt_path = @"..\..\phone.txt";
            sfz_path = @"..\..\out";
#endif


            //if (args.Length==0)
            //{
            //    Console.WriteLine("参数呢 \n  -h 查看帮助");

            //}
            //for (int arg=0;arg<args.Length;arg++)
            //{

            //    if(args[arg]=="-i")
            //    {
            //        sfz_path = args[arg + 1];
            //    }
            //    if (args[arg] == "-t")
            //    {
            //        txt_path = args[arg + 1];

            //    }
            //    if (args[arg] == "-e")
            //    {
            //        excel_path = args[arg + 1];


            //    }
            //    if (args[arg] == "-h")
            //    {
            //        Console.WriteLine(help);
            //    }

            //    //Console.ReadKey();

            //}
            if (sfz_path!=""&&excel_path!=""&&txt_path!="")
            {
                XLWorkbook wb = new XLWorkbook(excel_path);
                IXLWorksheet ws = wb.Worksheets.First();//获得第一个Sheet。
                Console.WriteLine(ws);
                string txt_name = Path.GetFileNameWithoutExtension(txt_path);
                string[] lines = File.ReadAllLines(txt_path, Encoding.UTF8);
                DirectoryInfo MyPath = new DirectoryInfo(sfz_path);
                //DirectoryInfo[] dics = MyPath.GetDirectories();
                FileInfo[] img_files = MyPath.GetFiles();
                Console.WriteLine(img_files[0]);
                string last_sfz = "";
                for (int line = 0; line < lines.Length; line++)
                {
                    
                    Class1.WriteCells(ws, line + 1, 3, lines[line]);
                    Class1.WriteCells(ws, line + 1, 4, txt_name);
                    Class1.WriteCells(ws, line + 1, 2, last_sfz = shaixuan_sfz(img_files[line].Name));
                    
                    Class1.WriteCells(ws, line + 1, 1, shaixuan_name(img_files[line].Name));
                    //重命名
                    //File.Move(img_files[line].FullName,lines[line]+"-"+shaixuan_zhengfan(img_files[line].Name));
                    
                    wb.Save();


                }

               
                //Class1.WriteCells(ws, 1, 1, "百度");
                //for (int i = 2; i < 10; i++)
                //{
                //    for (int k = 1; k < i; k++)
                //    {
                //        Class1.WriteCells(ws, k, i, "百度");
                //    }
                //}
               
            }
            
            //Console.ReadKey();

        }
    }
}
