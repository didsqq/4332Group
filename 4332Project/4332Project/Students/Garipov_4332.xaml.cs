using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Windows;
using Microsoft.Win32;
using Excel = Microsoft.Office.Interop.Excel;

namespace _4332Project.Students
{
    public partial class Garipov_4332 : Window
    {
        public Garipov_4332()
        {
            InitializeComponent();
        }
       

        private void Button_Click(object sender, RoutedEventArgs e)
        {
            OpenFileDialog ofd = new OpenFileDialog()
            {
                DefaultExt = "*.xls;*.xlsx",
                Filter = "���� Excel (Spisok.xlsx)|*.xlsx",
                Title = "�������� ���� ���� ������"
            };
            if (!(ofd.ShowDialog() == true))
                return;
            string[,] list;
            Excel.Application ObjWorkExcel = new Excel.Application();
            Excel.Workbook ObjWorkBook = ObjWorkExcel.Workbooks.Open(ofd.FileName);
            Excel.Worksheet ObjWorkSheet = (Excel.Worksheet)ObjWorkBook.Sheets[1];
            var lastCell = ObjWorkSheet.Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell);
            int _columns = (int)lastCell.Column;
            int _rows = (int)lastCell.Row;
            list = new string[_rows, _columns];
            for (int j = 0; j < _columns; j++)
            {
                for (int i = 0; i < _rows; i++)
                {
                    list[i, j] = ObjWorkSheet.Cells[i + 1, j + 1].Text;
                }
            }
            int lastRow = 0;
            for (int i = 0; i < _rows; i++)
            {
                if (list[i, 1] != string.Empty)
                {
                    lastRow = i;
                }
            }
            ObjWorkBook.Close(false, Type.Missing, Type.Missing);
            ObjWorkExcel.Quit();
            GC.Collect();

            using (isrpo2Entities usersEntities = new isrpo2Entities())
            {
                for (int i = 1; i <= lastRow; i++)
                {
                    var zakaz = new Services()
                    {
                        Id = Convert.ToInt32(list[i, 0]),
                        ServiceName = list[i, 1],
                        ServiceType = list[i, 2],
                        Price = Convert.ToDecimal(list[i, 4]),
                    };
                    usersEntities.Services.Add(zakaz);
                }
                usersEntities.SaveChanges();
            }
            MessageBox.Show("�������� �������������� ������", "�����", MessageBoxButton.OK, MessageBoxImage.Information);

        }

        private void Button_Click_1(object sender, RoutedEventArgs e)
        {
            // �������� ������ �� ���� ������
            List<Services> services;
            using (isrpo2Entities usersEntities = new isrpo2Entities())
            {
                services = usersEntities.Services.ToList();
            }

            // ���������� ������ �� ���������
            var category1 = services.Where(s => s.Price >= 0 && s.Price <= 350).ToList();
            var category2 = services.Where(s => s.Price > 350 && s.Price <= 800).ToList();
            var category3 = services.Where(s => s.Price > 800).ToList();

            // ������� ����� Excel-����
            var saveFileDialog = new SaveFileDialog
            {
                Filter = "Excel files (*.xlsx)|*.xlsx",
                Title = "��������� ���� Excel",
                FileName = "ExportedData.xlsx"
            };

            if (saveFileDialog.ShowDialog() == true)
            {
                Excel.Application excelApp = new Excel.Application();
                Excel.Workbook workbook = excelApp.Workbooks.Add();
                Excel.Worksheet worksheet;

                try
                {
                    // ���� ��� ��������� 1
                    worksheet = (Excel.Worksheet)workbook.Sheets[1]; // ������������� ����������
                    worksheet.Name = "��������� 1 (0-350)";
                    AddDataToWorksheet(worksheet, category1);

                    // ���� ��� ��������� 2
                    worksheet = (Excel.Worksheet)workbook.Sheets.Add(); // ������������� ����������
                    worksheet.Name = "��������� 2 (350-800)";
                    AddDataToWorksheet(worksheet, category2);

                    // ���� ��� ��������� 3
                    worksheet = (Excel.Worksheet)workbook.Sheets.Add(); // ������������� ����������
                    worksheet.Name = "��������� 3 (800+)";
                    AddDataToWorksheet(worksheet, category3);

                    // ��������� ����
                    workbook.SaveAs(saveFileDialog.FileName);
                    MessageBox.Show("������ ������� �������������� � Excel!", "�����", MessageBoxButton.OK, MessageBoxImage.Information);
                }
                catch (Exception ex)
                {
                    MessageBox.Show($"������ ��� �������� ������: {ex.Message}", "������", MessageBoxButton.OK, MessageBoxImage.Error);
                }
                finally
                {
                    // ��������� Excel
                    workbook.Close(false);
                    excelApp.Quit();
                    ReleaseObject(workbook);
                    ReleaseObject(excelApp);
                }
            }
        }

        private void AddDataToWorksheet(Excel.Worksheet worksheet, List<Services> data)
        {
            // ��������� ��������
            worksheet.Cells[1, 1] = "Id";
            worksheet.Cells[1, 2] = "�������� ������";
            worksheet.Cells[1, 3] = "��� ������";
            worksheet.Cells[1, 4] = "���������";

            // ��������� ������
            for (int i = 0; i < data.Count; i++)
            {
                worksheet.Cells[i + 2, 1] = data[i].Id;
                worksheet.Cells[i + 2, 2] = data[i].ServiceName;
                worksheet.Cells[i + 2, 3] = data[i].ServiceType;
                worksheet.Cells[i + 2, 4] = data[i].Price;
            }

            // ����-������ ��������
            worksheet.Columns.AutoFit();
        }

        private void ReleaseObject(object obj)
        {
            try
            {
                if (obj != null)
                {
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(obj);
                    obj = null;
                }
            }
            catch (Exception ex)
            {
                obj = null;
                MessageBox.Show($"������ ��� ������������ ��������: {ex.Message}", "������", MessageBoxButton.OK, MessageBoxImage.Error);
            }
            finally
            {
                GC.Collect();
            }
        }
    }
}