using Microsoft.Win32;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Windows;
using Excel = Microsoft.Office.Interop.Excel;
using System.Runtime.Serialization.Json;
using System.Text.Json;
using System.Security.Cryptography.X509Certificates;
using System.Security.Cryptography;
using System.Text.Json.Serialization;
using Microsoft.Office.Interop.Excel;
using System.Web.Script.Serialization;
using Newtonsoft.Json;
using Word = Microsoft.Office.Interop.Word;
using System.Data.Entity.Validation;
using System.ComponentModel;
using Newtonsoft.Json.Converters;

namespace Template_4337
{
    /// <summary>
    /// Логика взаимодействия для _4337_Gumerov.xaml
    /// </summary>
    public partial class _4337_Gumerov : System.Windows.Window
    {
        public _4337_Gumerov()
        {
            InitializeComponent();

        }

        private void Import(object sender, RoutedEventArgs e)
        {
            OpenFileDialog ofd = new OpenFileDialog()
            {
                DefaultExt = "*.xls;*.xlsx",
                Filter = "файл Excel (Spisok.xlsx)|*.xlsx",
                Title = "Выберите файл базы данных"
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
                for (int i = 0; i < _rows; i++)
                    list[i, j] = ObjWorkSheet.Cells[i + 1, j + 1].Text;
            ObjWorkBook.Close(false, Type.Missing, Type.Missing);
            ObjWorkExcel.Quit();
            GC.Collect();

            using (ISRPO2Entities usersEntities = new ISRPO2Entities())
            {
                for (int i = 0; i < _rows; i++)
                {
                    usersEntities.Clients.Add(new Clients()
                    {
                        FIO = list[i, 1],
                        Email = list[i, 2],
                        Age = list[i, 3]
                    });
                }
                usersEntities.SaveChanges();
            }

        }

        private void Export(object sender, RoutedEventArgs e)
        {
            List<Clients> allStudents;

            using (ISRPO2Entities usersEntities = new ISRPO2Entities())
            {
                allStudents = usersEntities.Clients.ToList().OrderBy(s => s.ClientCod).ToList();
            }
            var app = new Excel.Application();
            app.SheetsInNewWorkbook = allStudents.Count();
            Excel.Workbook workbook = app.Workbooks.Add(Type.Missing);

            for (int i = 1; i < 4; i++)
            {
                int startRowIndex = 1;
                Excel.Worksheet worksheet = app.Worksheets.Item[i];
                worksheet.Name = "Категория " + Convert.ToString(i);
                worksheet.Cells[1][startRowIndex] = "Код клиента";
                worksheet.Cells[2][startRowIndex] = "ФИО";
                worksheet.Cells[3][startRowIndex] = "Email";
                worksheet.Cells[4][startRowIndex] = "Возраст";
                startRowIndex++;

                foreach (var client in allStudents)
                {
                        string tip = "";
                        if (Convert.ToInt32(client.Age) <= 29 && Convert.ToInt32(client.Age) >= 20) { tip = "Категория 1"; }
                        if (Convert.ToInt32(client.Age) <= 39 && Convert.ToInt32(client.Age) >= 30) { tip = "Категория 2"; }
                        if (Convert.ToInt32(client.Age) >= 40) { tip = "Категория 3"; }
                        if (tip == worksheet.Name)
                        {
                            worksheet.Cells[1][startRowIndex] = client.ClientCod;
                            worksheet.Cells[2][startRowIndex] = client.FIO;
                            worksheet.Cells[3][startRowIndex] = client.Email;
                            worksheet.Cells[4][startRowIndex] = client.Age;
                            startRowIndex++;
                        }

                }

                worksheet.Columns.AutoFit();
            }
            app.Visible = true;

        }
        private void ImportJSON(object sender, RoutedEventArgs e)
        {
            using (ISRPO2Entities db = new ISRPO2Entities())
            {


                OpenFileDialog open_dialog = new OpenFileDialog();
                if (open_dialog.ShowDialog() == true)
                {
                        string json = File.ReadAllText(open_dialog.FileName);
                        json = json.Substring(0, json.Length - 1);
                        string[] words = json.Split('}');
                        string d = "";
                        foreach (string s in words)
                        {
                            d = s + "}";
                            d = d.Substring(1);
                            if (d != "")
                            {
                            var dateTimeConverter = new IsoDateTimeConverter { DateTimeFormat = "dd/MM/yyyy" };
                            ClientsJSON cl = JsonConvert.DeserializeObject<ClientsJSON>(d, dateTimeConverter);
                                db.ClientsJSON.Add(new ClientsJSON()
                                {
                                    Id = cl.Id,
                                    FullName = cl.FullName,
                                    E_Mail = cl.E_Mail,
                                    BirthDate = cl.BirthDate
                                });

                            }
                        }
                        try { 
                        db.SaveChanges();
                        MessageBox.Show("Успешно импортировано!");
                          }
                    catch (DbEntityValidationException ex)
                    {
                        foreach (DbEntityValidationResult validationError in ex.EntityValidationErrors)
                        {
                            MessageBox.Show("Object: " + validationError.Entry.Entity.ToString());
                            MessageBox.Show("\n");
                            foreach (DbValidationError err in validationError.ValidationErrors)
                            {
                                MessageBox.Show(err.ErrorMessage);
                                break;
                            }
                            break;
                        }
                    }
                }

            }

        }
    
        

        private void ExportWORD(object sender, RoutedEventArgs e)
        {
            List<ClientsJSON> allStudents;

            using (ISRPO2Entities usersEntities = new ISRPO2Entities())
            {
                allStudents = usersEntities.ClientsJSON.ToList();
            }
            var app = new Word.Application();
            Word.Document document = app.Documents.Add();

            for (int i = 1; i < 4; i++)
            {
                Word.Paragraph paragraph = document.Paragraphs.Add();
                Word.Range range = paragraph.Range;
                range.Text = "Категория " + Convert.ToString(i);
                string worksheet = "Категория " + Convert.ToString(i);
                paragraph.set_Style("Заголовок 1");
                range.InsertParagraphAfter();

                Word.Paragraph tableParagraph = document.Paragraphs.Add();
                Word.Range tableRange = tableParagraph.Range;
                Word.Table studentsTable = document.Tables.Add(tableRange, 15, 4);

                Word.Range cellRange;
                cellRange = studentsTable.Cell(1, 1).Range;
                cellRange.Text = "Код клиента";
                cellRange = studentsTable.Cell(1, 2).Range;
                cellRange.Text = "ФИО";
                cellRange = studentsTable.Cell(1, 3).Range;
                cellRange.Text = "Email";
                cellRange = studentsTable.Cell(1, 4).Range;
                cellRange.Text = "Возраст";
                studentsTable.Rows[1].Range.Bold = 1;
                studentsTable.Rows[1].Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;

                int k = 1;
                foreach (var clients in allStudents)
                {
                    string tip = "";
                    if (Convert.ToInt32(DateTime.Today.Year - clients.BirthDate.Value.Year) >= 20 && Convert.ToInt32(DateTime.Today.Year - clients.BirthDate.Value.Year) <= 29) { tip = "Категория 1"; }
                    if (Convert.ToInt32(DateTime.Today.Year - clients.BirthDate.Value.Year) >= 30 && Convert.ToInt32(DateTime.Today.Year - clients.BirthDate.Value.Year) <= 39) { tip = "Категория 2"; }
                    if (Convert.ToInt32(DateTime.Today.Year - clients.BirthDate.Value.Year) >= 40) { tip = "Категория 3"; }
                    if (tip == worksheet)
                    {
                        cellRange = studentsTable.Cell(k + 1, 1).Range;
                        cellRange.Text = clients.Id.ToString();
                        cellRange.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;
                        cellRange = studentsTable.Cell(k + 1, 2).Range;
                        cellRange.Text = clients.FullName;
                        cellRange.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;
                        cellRange = studentsTable.Cell(k + 1, 3).Range;
                        cellRange.Text = clients.E_Mail;
                        cellRange.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;
                        cellRange = studentsTable.Cell(k + 1, 4).Range;
                        cellRange.Text = (Convert.ToInt32(DateTime.Today.Year - clients.BirthDate.Value.Year).ToString());
                        cellRange.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;
                        k++;
                    }

                }
                document.Words.Last.InsertBreak(Word.WdBreakType.wdPageBreak);
            }
            app.Visible = true;
            MessageBox.Show("Документ сохранен");
            document.SaveAs2(@"C:\Users\Владелец\Desktop\outputFileWord.docx");
        }
    }

    }

