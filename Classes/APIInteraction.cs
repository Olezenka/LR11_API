using Newtonsoft.Json.Linq;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Net;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.Windows;
using Word = Microsoft.Office.Interop.Word;

namespace LR11.Classes
{
    internal class APIInteraction
    {
        private string fullname;

        private bool ContainsExtraChar(string input)
        {
            return !Regex.IsMatch(input, @"^[А-Яа-я\s]+$");
        }
        public string GetFullName()
        {
            string URL = "http://localhost:4444/TransferSimulator/";
            HttpWebRequest request = (HttpWebRequest)WebRequest.Create(URL);
            request.Method = "GET";

            HttpWebResponse response = (HttpWebResponse)request.GetResponse();

            StreamReader reader = new StreamReader(response.GetResponseStream());

            string text = reader.ReadToEnd();

            JObject jObject = JObject.Parse(text);

            string value = (string)jObject["value"];

            fullname = value;

            return fullname;
        }

        public string FillDocument()
        {
            string result = "";
            if (fullname == null)
            {
                MessageBox.Show("Данные не были получены");
                return result;
            }
            bool isValidFullName = ContainsExtraChar(fullname);
            result = isValidFullName ? "ФИО содержит запрещенные символы" : "ФИО не содержит запрещенные символы";

            string[] rowData = { $"Введите данные \n{fullname}", result, "Успешно" };

            AddToWordTable(rowData);
            return result;
        }
        public void AddToWordTable(string[] rowData)
        {
            var openFileDlg = new System.Windows.Forms.OpenFileDialog();
            string filePath;

            if (openFileDlg.ShowDialog() == System.Windows.Forms.DialogResult.OK)
                filePath = openFileDlg.FileName;
            else
                return;

            Word.Application wordApp = new Word.Application();
            Word.Document doc = null;

            try
            {
                doc = wordApp.Documents.Open(filePath);
                wordApp.Visible = false;

                Word.Table table = doc.Tables[1];
                Word.Row row = table.Rows.Add();
                for (int i = 0; i < rowData.Length; i++)
                {
                    row.Cells[i + 1].Range.Text = rowData[i]
;
                }
                doc.Save();
                MessageBox.Show("Информация была добавлена в файлы!");
            }
            catch
            {
                MessageBox.Show("Ошибка при работе с документом");

            }
            finally
            {
                if (doc != null)
                {
                    doc.Close(Word.WdSaveOptions.wdSaveChanges);
                }
                wordApp.Quit(Word.WdSaveOptions.wdSaveChanges);
            }
        }
    }
}
