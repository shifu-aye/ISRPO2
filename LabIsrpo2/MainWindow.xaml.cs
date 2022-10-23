using Microsoft.Office.Interop.Excel;
using Microsoft.Win32;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows;
using System.Windows.Documents;
using Excel = Microsoft.Office.Interop.Excel;
using Window = System.Windows.Window;

namespace LabIsrpo2
{
    /// <summary>
    /// Логика взаимодействия для MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        public MainWindow()
        {
            InitializeComponent();
        }

        /// <summary>
        /// Обработчик событий для импорта
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void ImportBtn_OnClick(object sender, RoutedEventArgs e)
        {
            OpenFileDialog fd = new OpenFileDialog()
            {
                DefaultExt = "*.xls;*.xlsx",
                Filter = "файл Excel (Spisok.xlsx)|*.xlsx",
                Title = "Выберите файл"
            };
            if (!(fd.ShowDialog()) == true)
            {
                return;
            }

            Import(fd.FileName);
        }

        /// <summary>
        /// Метод импорта данных
        /// </summary>
        /// <param name="fn">Имя файла Excel</param>
        private void Import(string fn)
        {
            try
            {
                string[,] list;
                Excel.Application ObjWorkExcel = new Excel.Application();
                Excel.Workbook ObjWorkBook = ObjWorkExcel.Workbooks.Open(fn);
                Excel.Worksheet ObjWorkSheet = (Excel.Worksheet)ObjWorkBook.Sheets[1];
                var lastCell = ObjWorkSheet.Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell);
                int _columns = (int)lastCell.Column;
                int _rows = (int)lastCell.Row;
                list = new string[_rows, _columns];
                for (int j = 0; j < _columns; j++)
                {
                    for (int i = 1; i < _rows; i++)
                    {
                        list[i, j] = ObjWorkSheet.Cells[i + 1, j + 1].Text;
                    }
                }

                ObjWorkBook.Close(false, Type.Missing, Type.Missing);
                ObjWorkExcel.Quit();
                GC.Collect();
                using (ServiceEntities db = new ServiceEntities())
                {
                    for (int i = 1; i < _rows; i++)
                    {
                        db.ServicesTable.Add(new ServicesTable()
                        {
                            OrderCode = list[i, 1],
                            DateOfCreation = list[i, 2],
                            OrderTIme = list[i, 3],
                            ClientCode = list[i, 4],
                            Services = list[i, 5],
                            Status = list[i, 6],
                            DateClosed = list[i, 7],
                            RollerTime = list[i, 8]
                        });
                    }

                    db.SaveChanges();
                    ExcelDg.ItemsSource = db.ServicesTable.ToList();
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }

        }

        /// <summary>
        /// Очистка базы данных
        /// </summary>
        private void RemoveData()
        {
            using (ServiceEntities db = new ServiceEntities())
            {
                db.ServicesTable.RemoveRange(db.ServicesTable.ToList());
                db.SaveChanges();
            }
        }

        /// <summary>
        /// Обработчик событий для импорта
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        /// <exception cref="NotImplementedException"></exception>
        private void ExportBtn_OnClick(object sender, RoutedEventArgs e)
        {
            Export();
        }
        /// <summary>
        /// Метод экспорта
        /// </summary>
        private void Export()
        {
            using (ServiceEntities db = new ServiceEntities())
            {
                var service = db.ServicesTable.GroupBy(p => p.DateOfCreation).ToList();
                
                var application = new Excel.Application();
                application.SheetsInNewWorkbook = service.Count();
                Excel.Workbook workbook = application.Workbooks.Add(Type.Missing);

                int j = 1;
                foreach (var s in service)
                {
                    Excel.Worksheet worksheet = application.Worksheets.Item[j];
                    worksheet.Name = s.Key;
                    j++;
                }

                for (int i = 0; i < application.SheetsInNewWorkbook; i++)
                {
                    int startRowIndex = 1;

                    Excel.Worksheet worksheet = application.Worksheets.Item[i + 1];

                    worksheet.Cells[1][2] = "Порядковый номер";
                    worksheet.Cells[2][2] = "Код заказа";
                    worksheet.Cells[3][2] = "Код клиента";
                    worksheet.Cells[4][2] = "Услуги";

                    startRowIndex++;

                    foreach (var data in service)
                    {
                        if (data.Key == Convert.ToString(worksheet.Name))
                        {
                            Excel.Range headerRange = worksheet.Range[worksheet.Cells[1][1], worksheet.Cells[2][1]];
                            headerRange.Merge();
                            headerRange.Value = Convert.ToString(worksheet.Name);
                            headerRange.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                            headerRange.Font.Italic = true;

                            startRowIndex++;

                            foreach (ServicesTable services in db.ServicesTable)
                            {
                                if (services.DateOfCreation == data.Key)
                                {
                                    worksheet.Cells[1][startRowIndex] = services.ID;
                                    worksheet.Cells[2][startRowIndex] = services.OrderCode;
                                    worksheet.Cells[3][startRowIndex] = services.ClientCode;
                                    worksheet.Cells[4][startRowIndex] = services.Services;

                                    startRowIndex++;
                                }
                            }
                        }
                        else
                        {
                            continue;
                        }
                    }

                    Excel.Range rangeBorders =
                        worksheet.Range[worksheet.Cells[1][1], worksheet.Cells[4][startRowIndex - 1]];
                    rangeBorders.Borders[Excel.XlBordersIndex.xlEdgeBottom].LineStyle =
                        rangeBorders.Borders[Excel.XlBordersIndex.xlEdgeLeft].LineStyle =
                            rangeBorders.Borders[Excel.XlBordersIndex.xlEdgeTop].LineStyle =
                                rangeBorders.Borders[Excel.XlBordersIndex.xlEdgeRight].LineStyle =
                                    rangeBorders.Borders[Excel.XlBordersIndex.xlInsideHorizontal].LineStyle =
                                        rangeBorders.Borders[Excel.XlBordersIndex.xlInsideVertical].LineStyle =
                                            Excel.XlLineStyle.xlContinuous;

                    worksheet.Columns.AutoFit();
                }

                application.Visible = true;
            }
        }

        private void Window_Closed(object sender, EventArgs e)
        {
            RemoveData();
        }
    }
}