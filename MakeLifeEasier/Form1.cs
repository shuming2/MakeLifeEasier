using System;
using System.IO;
using System.Windows.Forms;

using NPOI.SS.UserModel;
using NPOI.HSSF.UserModel;
using NPOI.XSSF.UserModel;
using System.Collections.Generic;
using System.Collections;

namespace MakeLifeEasier
{
    public partial class Form1 : Form
    {
        private string inputExcelPath, outputExcelPath;

        public Form1()
        {
            InitializeComponent();
        }

        private void BrowseInputButton_Click(object sender, EventArgs e)
        {
            //openFileDialog1.InitialDirectory = Environment.GetFolderPath(Environment.SpecialFolder.Desktop);
            openFileDialog1.Filter = "Excel文件|*.xls;*.xlsx";
                                                        
            if (this.openFileDialog1.ShowDialog() == DialogResult.OK)
            {
                inputExcelPath = openFileDialog1.FileName;
                InputExcelPathLabel.Text = Path.GetFileName(inputExcelPath);
            }
        }

        /*
        private void BrowseOutputButton_Click(object sender, EventArgs e)
        {
            //openFileDialog2.InitialDirectory = Environment.GetFolderPath(Environment.SpecialFolder.Desktop);
            openFileDialog2.Filter = "Excel文件|*.xls;*.xlsx";

            if (this.openFileDialog2.ShowDialog() == DialogResult.OK)
            {
                outputExcelPath = openFileDialog2.FileName;
                OutputExcelPathLabel.Text = Path.GetFileName(outputExcelPath);
            }
        }
        */

        private List<List<String>> GetRestDateofThisMonth()
        {
            DateTime startDate = this.dateTimePicker1.Value;
            List<List<String>> dateRange = new List<List<String>>();

            DateTime d1 = new DateTime(startDate.Year, startDate.Month, 1);
            DateTime endDate = d1.AddMonths(1).AddDays(-1);
            TimeSpan span = endDate.Subtract(startDate);
            int dayDiff = span.Days + 1;
            if (startDate.Day == endDate.Day)
            {
                dayDiff = 0;
            }
            DateTime newDate = new DateTime();
            for (int i = 0; i <= dayDiff; i++)
            {
                newDate = startDate.AddDays(i);
                string[] dateStrl = { newDate.ToString("yyyy/M/d 0:00"), newDate.ToString("yyyy/M/d 12:00") };
                dateRange.Add(new List<string>(dateStrl));
                Console.WriteLine(dateStrl[0]);
            }
            return dateRange;
        }

        private List<List<String>> GetInputData()
        {
            FileStream fileStream = new FileStream(inputExcelPath, FileMode.Open, FileAccess.Read);

            IWorkbook inputExcel = null;
            if (inputExcelPath.IndexOf(".xlsx") > 0) // 2007
            {
                inputExcel = new XSSFWorkbook(fileStream);
            }
            else if (inputExcelPath.IndexOf(".xls") > 0) // 2003
            {
                inputExcel = new HSSFWorkbook(fileStream);
            }

            ISheet sheet = inputExcel.GetSheetAt(0);
            IRow row;

            // Get the column number of city, town, district
            row = sheet.GetRow(0);
            int columnNum = row.LastCellNum;
            List<String> titles = new List<String>();
            for (int i = 0; i < columnNum; i++)
            {
                titles.Add(row.GetCell(i).ToString());
            }
            int cityIndex = titles.FindIndex(a=>a=="地市");
            int townIndex = titles.FindIndex(a => a == "区县");
            int districtIndex = titles.FindIndex(a => a == "小区");
   

            List<List<String>> data = new List<List<String>>();
            for (int i = 1; i < sheet.LastRowNum; i++)  // Every row except the first
            {
                row = sheet.GetRow(i);
                if (row != null)
                {
                    string[] rowData = { row.GetCell(cityIndex).ToString(), row.GetCell(townIndex).ToString(), row.GetCell(districtIndex).ToString() };
                    data.Add(new List<String>(rowData));

                }
            }

            fileStream.Close();
            inputExcel.Close();

            return data;
        }

        private List<List<String>> ConvertData(List<List<String>> data, List<List<String>> dateRange)
        {
            List<List<String>> newData = new List<List<String>>();
            foreach (List<String> date in dateRange)
            {
                foreach (List<String> t in data)
                {
                    String[] rowData ={ "聚类市场断电", "工程", date[0], date[1], t[2], t[0], t[1] };
                    newData.Add(new List<String>(rowData));
                }
            }
            return newData;
        }

        private List<List<String>> GetOutputData()
        {
            List<List<String>> data = new List<List<String>>();

            if (outputExcelPath == null)
            {
                return data;
            }

            FileStream fileStream = new FileStream(outputExcelPath, FileMode.Open, FileAccess.Read);

            IWorkbook outputExcel = null;
            if (outputExcelPath.IndexOf(".xlsx") > 0) // 2007
            {
                outputExcel = new XSSFWorkbook(fileStream);
            }
            else if (outputExcelPath.IndexOf(".xls") > 0) // 2003
            {
                outputExcel = new HSSFWorkbook(fileStream);
            }

            ISheet sheet = outputExcel.GetSheetAt(0);
            IRow row;
            ICell cell;

            for (int i = 1; i < sheet.LastRowNum; i++)  // Every row except the first
            {
                row = sheet.GetRow(i);
                if (row != null)
                {
                    List<String> rowData = new List<string>();
                    for (int j = 0; j < 7; j++)
                    {
                        cell = row.GetCell(j);
                        if (cell.CellType == CellType.String)
                        {
                            rowData.Add(cell.ToString());
                        }
                        else
                        {
                            rowData.Add(cell.DateCellValue.ToString("yyyy/m/d h:mm"));
                        }
                        
                    }
                    data.Add(new List<String>(rowData));
                }
            }

            fileStream.Close();
            outputExcel.Close();

            return data;
        }

        private void WriteToNewFile(List<List<String>> data)
        {
            XSSFWorkbook newWorkbook = new XSSFWorkbook();  
            newWorkbook.CreateSheet("Sheet1");
            ISheet sheet = newWorkbook.GetSheetAt(0);
            sheet.SetColumnWidth(2, 20 * 256);
            sheet.SetColumnWidth(3, 20 * 256);

            sheet.CreateRow(0);
            IRow row = sheet.GetRow(0);
            List<ICell> cells = new List<ICell>();
            for (int j = 0; j < 7; j++)
            {
                cells.Add(row.CreateCell(j));
            }
            cells[0].SetCellValue("割接名称");
            cells[1].SetCellValue("任务类型");
            cells[2].SetCellValue("开始时间");
            cells[3].SetCellValue("结束时间");
            cells[4].SetCellValue("小区名称");
            cells[5].SetCellValue("地市");
            cells[6].SetCellValue("区县");

            IDataFormat dataformat = newWorkbook.CreateDataFormat();
            ICellStyle style0 = newWorkbook.CreateCellStyle();
            style0.DataFormat = dataformat.GetFormat("yyyy/m/d h:mm");

            for (int i = 1; i < data.Count + 1; i++)
            {
                sheet.CreateRow(i);
                row = sheet.GetRow(i);
                for (int j = 0; j < 7; j++)
                {
                    ICell cell = row.CreateCell(j);
                    if (j == 2 || j == 3)
                    {
                        cell.SetCellValue(Convert.ToDateTime(data[i - 1][j]));
                        cell.CellStyle = style0;
                    }else{
                        cell.SetCellValue(data[i - 1][j]);
                    }
                }
            }

            FileStream newFile = new FileStream(Environment.GetFolderPath(Environment.SpecialFolder.Desktop)+ @"\MakeLifeEasier.xlsx", FileMode.Create);
            newWorkbook.Write(newFile);
            newFile.Close();
            newWorkbook.Close();
        }

        private void StartButton_Click(object sender, EventArgs e)
        {
            if (inputExcelPath == null)
            {
                return;
            }
            List<List<String>> dateRange = this.GetRestDateofThisMonth();

            List<List<String>> inputData = this.GetInputData();

            List<List<String>> newData = this.ConvertData(inputData, dateRange);

            List<List<String>> dataToMerge = this.GetOutputData();

            dataToMerge.AddRange(newData);

            WriteToNewFile(dataToMerge);

            MessageBox.Show("MakeLifeEasier.xlsx 已在桌面生成", "生成成功");
        }

    }
}
