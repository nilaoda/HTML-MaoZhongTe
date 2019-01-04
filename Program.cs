using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using System.Collections;
using System.IO;
using System.Text;

namespace 毛概Excel处理
{
    class Program
    {
        /// <summary>
        /// 寻找指定目录下指定后缀的文件的详细路径 如".txt"
        /// </summary>
        /// <param name="dir"></param>
        /// <param name="ext"></param>
        /// <returns></returns>
        public static string[] GetFiles(string dir, string ext)
        {
            ArrayList al = new ArrayList();
            StringBuilder sb = new StringBuilder();
            DirectoryInfo d = new DirectoryInfo(dir);
            foreach (FileInfo fi in d.GetFiles())
            {
                if (fi.Extension.ToUpper() == ext.ToUpper())
                {
                    al.Add(fi.FullName);
                }
            }
            return (string[])al.ToArray(typeof(string));
        }

        static void Main(string[] args)
        {
            foreach (string ff in GetFiles(@"D:\毛概题库", ".xls"))
            {
                string filePath = ff;
                StringBuilder sb = new StringBuilder();
                FileStream fs = null;
                IWorkbook workbook = null;
                ISheet sheet = null;

                using (fs = File.OpenRead(filePath))
                {
                    // 2007版本  
                    if (filePath.IndexOf(".xlsx") > 0)
                        workbook = new XSSFWorkbook(fs);
                    // 2003版本  
                    else if (filePath.IndexOf(".xls") > 0)
                        workbook = new HSSFWorkbook(fs);

                    if (workbook != null)
                    {
                        sheet = workbook.GetSheetAt(0);//读取第一个sheet
                        if (sheet != null)
                        {
                            int rowCount = sheet.LastRowNum;//总行数  
                            if (rowCount > 0)
                            {
                                for (int i = 1; i <= rowCount; i++)
                                {
                                    string daan = string.Empty;
                                    string current = sheet.GetRow(i).GetCell(3).ToString(); //ABCD
                                    for (int ii = 0; ii < current.Length; ii++)
                                    {
                                        if (sheet.GetRow(i).GetCell(1).ToString() == "简答题")
                                        {
                                            daan += sheet.GetRow(i).GetCell(7).ToString();
                                            break;
                                        }
                                        if (sheet.GetRow(i).GetCell(1).ToString() == "判断题")
                                        {
                                            if (current[ii] == 'A')
                                                daan += "正确";
                                            else if (current[ii] == 'B')
                                                daan += "错误";
                                            break;
                                        }
                                        if (sheet.GetRow(i).GetCell(1).ToString() == "填空题")
                                        {
                                            if (current[ii] == 'A' && sheet.GetRow(i).GetCell(7) != null)
                                                daan += sheet.GetRow(i).GetCell(7).ToString() + "\r\n";
                                            else if (current[ii] == 'B' && sheet.GetRow(i).GetCell(8) != null)
                                                daan += sheet.GetRow(i).GetCell(8).ToString() + "\r\n";
                                            else if (current[ii] == 'C' && sheet.GetRow(i).GetCell(9) != null)
                                                daan += sheet.GetRow(i).GetCell(9).ToString() + "\r\n";
                                            else if (current[ii] == 'D' && sheet.GetRow(i).GetCell(10) != null)
                                                daan += sheet.GetRow(i).GetCell(10).ToString() + "\r\n";
                                            break;
                                        }
                                        if (sheet.GetRow(i).GetCell(1).ToString() == "单选题"
                                            && current[ii] == 'D' && sheet.GetRow(i).GetCell(10).ToString() == "以上都对")
                                        {
                                            daan += "A." + sheet.GetRow(i).GetCell(7).ToString() + "\r\n";
                                            daan += "B." + sheet.GetRow(i).GetCell(8).ToString() + "\r\n";
                                            daan += "C." + sheet.GetRow(i).GetCell(9).ToString() + "\r\n";
                                            break;
                                        }
                                        if (current[ii] == 'A' && sheet.GetRow(i).GetCell(7) != null)
                                            daan += "A." + sheet.GetRow(i).GetCell(7).ToString() + "\r\n";
                                        else if (current[ii] == 'B' && sheet.GetRow(i).GetCell(8) != null)
                                            daan += "B." + sheet.GetRow(i).GetCell(8).ToString() + "\r\n";
                                        else if (current[ii] == 'C' && sheet.GetRow(i).GetCell(9) != null)
                                            daan += "C." + sheet.GetRow(i).GetCell(9).ToString() + "\r\n";
                                        else if (current[ii] == 'D' && sheet.GetRow(i).GetCell(10) != null)
                                            daan += "D." + sheet.GetRow(i).GetCell(10).ToString() + "\r\n";
                                    }
                                    //string dd = i + "." + sheet.GetRow(i).GetCell(2).ToString() + "\r\n" + daan.Trim() + "\r\n";
                                    //sb.AppendLine(dd);
                                    string dd = "<p>" + i.ToString("0000") + "：" + sheet.GetRow(i).GetCell(2).ToString() + "</p>";
                                    dd += "<p>答案：" + daan.Trim() + "</p>";
                                    sb.Append(dd.Replace("\n", "<br>").Replace("\r", ""));
                                }
                            }
                        }
                    }
                }

                //File.WriteAllText(Path.GetDirectoryName(filePath) + "\\" + Path.GetFileNameWithoutExtension(filePath) + ".txt", sb.ToString());
                File.WriteAllText(Path.GetDirectoryName(filePath) + "\\" + Path.GetFileNameWithoutExtension(filePath) + ".html",
                    @"<html><head><meta name=""viewport"" content=""width=device-width, initial-scale=1.0, maximum-scale=1.0, minimum-scale=1.0, user-scalable=no"" /><style>.content {max-width: 800px;margin: auto;text-align: left;}p {text-indent: -3em;padding-left: 3em;}.content p {margin: .5em 0;}.content p:nth-of-type(even) {margin: .5em 0 1em;}</style></head><body>"
                    + "<h2 style=\"text-align:center\">" + Path.GetFileNameWithoutExtension(filePath) + "</h2><div class=\"content\">" + sb.ToString() + "</div></body><footer style=\"text-align:center\">由 N 整理</footer></html> ", Encoding.UTF8);
                //Console.ReadLine();
            }
        }
    }
}
