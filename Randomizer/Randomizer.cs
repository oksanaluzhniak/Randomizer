using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Windows.Forms;
using OfficeOpenXml;
using DocumentFormat.OpenXml.Packaging;
using text = DocumentFormat.OpenXml.Wordprocessing.Text;


namespace Randomizer
{
    public partial class Randomizer : Form
    {
        public Randomizer()
        {
            InitializeComponent();
        }

        private void button2_Click(object sender, EventArgs e)
        {
            string fullName = (string)textBox1.Text + " " + (string)textBox2.Text;
            
            // Перевірка, чи є студент з введеним прізвищем
            string studentsExcelFilePath = "D:\\Diplom\\StudentsList.xlsx"; // Шлях до файлу зі списком студентів
            bool studentExists = CheckStudentExists(fullName, studentsExcelFilePath);
            if (studentExists)
            {
                int labNumber = int.Parse(textBox3.Text);

                // Визначення, чи потрібні студенту складні завдання
                bool isComplexTasks = IsComplexTasks(fullName, studentsExcelFilePath);

                // Вибір файлу з завданнями в залежності від складності завдань
                string templateFilePath;
                if (labNumber==1)
                {
                    // Заміна поля {StudentName} на значення studentName
                     templateFilePath = "D:\\Diplom\\Template1.docx";
                }
                else if (labNumber == 2)
                {
                    // Заміна поля {StudentName} на значення studentName
                    templateFilePath = "D:\\Diplom\\Template2.docx";
                }
                else
                {
                    // Заміна поля {StudentName} на значення studentName
                    templateFilePath = "D:\\Diplom\\Template3.docx";
                }
                string tasksTextFilePath;
                if (isComplexTasks)
                    tasksTextFilePath = "D:\\Diplom\\АтрибутиСкладні.txt"; // Шлях до файлу зі складними завданнями
                else
                    tasksTextFilePath = "D:\\Diplom\\АтрибутиПрості.txt"; // Шлях до файлу зі простими завданнями

                string [] tasks= SelectTaskForStudent(fullName, labNumber, tasksTextFilePath);
                string newFilePath = InsertValuesIntoTemplate(templateFilePath, fullName, labNumber,tasks);

                MessageBox.Show("Новий файл успішно створено за шаблоном: " + newFilePath);
            }
            else
            {
                MessageBox.Show("Студента з введеним прізвищем не знайдено.");
            }

        }

        static string InsertValuesIntoTemplate(string templateFilePath, string studentName, int labNumber, string[] tasks)
        {
            string newFilePath = $"D:\\Diplom\\{studentName} {labNumber}.docx"; // Шлях до нового файлу

            // Копіювання шаблонного файлу
            System.IO.File.Copy(templateFilePath, newFilePath, true);

            using (WordprocessingDocument doc = WordprocessingDocument.Open(newFilePath, true))
            {
                // Зберігаємо всі тексти в окремому списку
                var texts = doc.MainDocumentPart.Document.Body.Descendants<text>().ToList();

                // Проходимося по копії списку і замінюємо значення
                foreach (var text in texts)
                {
                    // Пошук і заміна відповідних полів
                    if (text.Text.Contains("fullName"))
                    {
                        text.Text = text.Text.Replace("fullName", studentName);
                    }
                    else if (text.Text.Contains("LabNumb"))
                    {
                        text.Text = text.Text.Replace("LabNumb", labNumber.ToString());
                    }
                    else if (text.Text.Contains("first"))
                    {
                        text.Text = text.Text.Replace("first", tasks[0]);
                    }
                    else if (text.Text.Contains("second"))
                    {
                        text.Text = text.Text.Replace("second", tasks[1]);
                    }
                    else if (text.Text.Contains("third"))
                    {
                        text.Text = text.Text.Replace("third", tasks[2]);

                    }
                    else if (text.Text.Contains("fourth"))
                    {
                        text.Text = text.Text.Replace("fourth", tasks[3]);
                    }
                }

                // Зберігаємо зміни у документ
                doc.MainDocumentPart.Document.Save();
            }
            

            return newFilePath;

        }

        // Перевірка, чи є студент з введеним прізвищем в файлі
        static bool CheckStudentExists(string fullName, string studentsExcelFilePath)
        {
            using (var package = new ExcelPackage(new FileInfo(studentsExcelFilePath)))
            {
                var worksheet = package.Workbook.Worksheets.FirstOrDefault();
                if (worksheet != null)
                {
                    int rowCount = worksheet.Dimension.Rows;

                    for (int i = 1; i <= rowCount; i++)
                    {
                        if (worksheet.Cells[i, 1].Value.ToString() == fullName)
                        {
                            return true;
                        }
                    }
                }
                else
                {
                    MessageBox.Show("У файлі немає листів.");
                }
            }

            return false;
        }

        // Метод для перевірки, чи потрібні студенту складні завдання
        static bool IsComplexTasks(string fullName, string studentsExcelFilePath)
        {
            using (var package = new ExcelPackage(new FileInfo(studentsExcelFilePath)))
            {
                var worksheet = package.Workbook.Worksheets.FirstOrDefault();
                if (worksheet != null)
                {
                    int rowCount = worksheet.Dimension.Rows;

                    for (int i = 1; i <= rowCount; i++)
                    {
                        if (worksheet.Cells[i, 1].Value.ToString() == fullName)
                        {
                            string taskType = worksheet.Cells[i, 2].Value.ToString();
                            return taskType == "+"; // Повертаємо true, якщо студентові потрібні складні завдання
                        }
                    }
                }
                else
                {
                    MessageBox.Show("У файлі немає листів.");
                }
            }

            return false; // Повертаємо false, якщо студентові потрібні прості завдання або не знайдено студента
        }

        // Метод для вибору завдання для студента
        static string[] ReadTasks(string tasksTextFilePath)
        {
            var lines = File.ReadAllLines(tasksTextFilePath);

            // Перевірка, чи є рядки для читання
            if (lines.Length == 0)
            {
                MessageBox.Show("Файл містить порожню або відсутню інформацію.");
                return new string[0]; // Повертаємо порожній масив завдань
            }


            // Вибірка завдань з рядків
            var tasks = new List<string>();
            foreach (var line in lines)
            {
                // Пропускаємо порожні рядки
                if (string.IsNullOrWhiteSpace(line))
                {
                    continue;
                }

                // Додаємо завдання до списку
                tasks.Add(line.Trim());
            }

            return tasks.ToArray();
        }

        static string[] SelectTaskForStudent(string fullName, int labNumber, string tasksTextFilePath)
        {
            // Зчитуємо завдання з файлу
            var tasks = ReadTasks(tasksTextFilePath);

            // Перевіряємо, чи є завдання для вказаної лабораторної роботи
            if (tasks.Length == 0)
            {
                MessageBox.Show($"Не вдалося знайти завдання для лабораторної роботи {labNumber}.");
                string[] strings = new string[0] ;
                return strings;
            }

            // Визначаємо початковий і кінцевий індекси рядків для поточної лабораторної роботи
            int startLineIndex = Array.IndexOf(tasks, $"Лабораторна робота {labNumber}") + 1;
            int endLineIndex = tasks.Length;
            if (labNumber < 3)
            {
                endLineIndex = Array.IndexOf(tasks, $"Лабораторна робота {labNumber + 1}");
            }

            // Створюємо словник для зберігання завдань з кожної групи
            var groupTasks = new Dictionary<string, List<string>>();

            // Проходимо по рядках з завданнями в межах поточної лабораторної роботи
            string currentGroup = "";
            for (int i = startLineIndex; i < endLineIndex; i++)
            {
                string line = tasks[i];

                if (line.StartsWith("Група"))
                {
                    // Оновлюємо поточну групу
                    currentGroup = line.Trim();
                    if (!groupTasks.ContainsKey(currentGroup))
                    {
                        groupTasks[currentGroup] = new List<string>();
                    }
                }
                else if (!string.IsNullOrWhiteSpace(line) && !line.StartsWith("Лабораторна робота"))
                {
                    // Додаємо завдання до поточної групи, розділяючи їх по комі
                    var tasksArray = line.Split(new[] { ',' }, StringSplitOptions.RemoveEmptyEntries)
                                          .Select(task => task.Trim())
                                          .ToList();
                    groupTasks[currentGroup].AddRange(tasksArray);
                }
            }

            // Вибираємо випадкове завдання з кожної групи
            var random = new Random();
            var selectedTasks = new List<string>();
            foreach (var tasksInGroup in groupTasks.Values)
            {
                if (tasksInGroup.Count > 0)
                {
                    int randomIndex = random.Next(tasksInGroup.Count);
                    selectedTasks.Add(tasksInGroup[randomIndex]);
                }
                else
                {
                    MessageBox.Show($"У групі {currentGroup} відсутні завдання.");
                }
            }

            // Виводимо обрані випадкові завдання для студента
            return selectedTasks.ToArray();
        }

        private void textBox2_TextChanged(object sender, EventArgs e)
        {
           
        }

        private void button1_Click(object sender, EventArgs e)
        {
            Close();
        }
    }
}
