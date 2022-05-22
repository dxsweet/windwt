using System.Diagnostics;
using System.Text;
using Excel1 = Microsoft.Office.Interop.Excel;


internal class Program
{
    static void Main(String[] args)
    {


        foreach(Process clsProcess in Process.GetProcesses())
        {
            if (clsProcess.ProcessName.Equals("EXCEL"))
            {
                clsProcess.Kill();
            }
        }

        Excel1.Application xlApp = new Excel1.Application();
        xlApp.Visible = false;
        xlApp.DisplayAlerts = false;

        DirectoryInfo di = new DirectoryInfo(@"./input");


        foreach (FileInfo fi in di.GetFiles())
        {
            Excel1.Workbook wb = xlApp.Workbooks.Open(fi.FullName);
            Excel1.Worksheet ws = wb.Worksheets["风向风速"];

            StringBuilder sb1 = new StringBuilder();
            //Cells.Item(Row, Column)
            StringBuilder sb2 = new StringBuilder();

            for(int i = 6; i <= 15; i++)
            {
                for(int j = 2; j <= 48; j += 2)
                {
                    sb1.AppendLine(ws.Cells[i,j].Value2);
                    sb2.AppendLine(ws.Cells[i,j+1].Value2);
                }
            }

            for (int i = 18; i <= 27; i++)
            {
                for (int j = 2; j <= 48; j += 2)
                {
                    sb1.AppendLine(ws.Cells[i, j].Value2);
                    sb2.AppendLine(ws.Cells[i, j+1].Value2);
                }
            }


            for (int i = 30; i <= 40; i++)
            {
                for (int j = 2; j <= 48; j += 2)
                {
                    sb1.AppendLine(ws.Cells[i, j].Value2);
                    sb2.AppendLine(ws.Cells[i, j+1].Value2);
                }
            }

            String s1 = sb1.ToString().Trim();
            sb1.Clear();
            String s2 = sb2.ToString().Trim();
            sb2.Clear();

            wb.Close();
            xlApp.Workbooks.Close();

            String fnfn1 = @"./output/" + fi.Name.Remove(7).Remove(4, 1) + "_1.txt";
            String fnfn2 = @"./output/" + fi.Name.Remove(7).Remove(4, 1) + "_2.txt";
            System.IO.File.WriteAllText(fnfn1, s1, Encoding.UTF8);
            System.IO.File.WriteAllText(fnfn2, s2, Encoding.UTF8);

            Console.WriteLine(fi.Name + "处理完毕");
        }

        xlApp.Quit();
        Console.WriteLine("全部处理完毕");

        foreach(Process clsProcess in Process.GetProcesses())
        {
            if (clsProcess.Equals("EXCEL"))
            {
                clsProcess.Kill();
            }
        }

        Console.ReadLine();
    }
}