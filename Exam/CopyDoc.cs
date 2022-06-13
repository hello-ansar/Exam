using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Word = Microsoft.Office.Interop.Word;
using System.Data.SqlClient;
using System.Data.Sql;
namespace Exam
{
    class CopyDoc
    {
        private FileInfo _fileInfo;

        public CopyDoc(string fileName)
        {
            if (File.Exists(fileName))
            {
                _fileInfo = new FileInfo(fileName);
            }
            else
            {
                throw new ArgumentException("File not found");
            }
        }

        internal bool Process(string text, Dictionary<string, string> items)
        {

            //List<String> que = new List<String>();

            Word.Application app = null;
            try
            {
                app = new Word.Application();
                Object file = _fileInfo.FullName;

                Object missing = Type.Missing;

                app.Documents.Open(file);

                int num = Convert.ToInt32(text);

                string copyText = "";
                for (int j = 0; j < app.ActiveDocument.Paragraphs.Count; j++)
                {
                    copyText +=  app.ActiveDocument.Paragraphs[j + 1].Range.Text;
                }


                app.ActiveDocument.Close();

                Object oEndOfDoc = "\\endofdoc";

                Word.Paragraph oPara2;

                SqlConnection sqlConnection = new SqlConnection("Data Source=DESKTOP-8NPMOGM\\SQLEXPRESS;Initial Catalog=Exam;Integrated Security=True");
                sqlConnection.Open();

                string request = "SELECT TOP (1000) questions FROM[Exam].[dbo].[Quest]";

                SqlCommand sqlCommand = new SqlCommand(request, sqlConnection);

                SqlDataReader sqlDataReader = sqlCommand.ExecuteReader();

                string[] mas = new string[20];

                int q = 0;
                while (sqlDataReader.Read())
                {
                    mas[q] = (string)sqlDataReader[0];
                    q++;
                }

                Console.WriteLine(sqlDataReader.HasRows);

                Random rnd = new Random();

                int q1, q2, q3;

                

                for (int i = 1; i < num; i++)
                {
                    var helper = new Helper("C:/Users/Ансар/source/repos/Exam/Exam/шаблон.docx");

                    items["<NUM>"] = i.ToString();

                    while (true)
                    {
                        q1 = (int)rnd.Next(0, 19);
                        q2 = (int)rnd.Next(0, 19);
                        q3 = (int)rnd.Next(0, 19);
                        if (q1 != q2 &&
                            q2 != q3 &&
                            q1 != q3)
                        {
                            break;
                        }
                    }

                    items["<Q1>"] = mas[q1];
                    items["<Q2>"] = mas[q2];
                    items["<Q3>"] = mas[q3];

                    helper.Process(items);

                    app.Documents.Open(file);
                    Object oRng = app.ActiveDocument.Bookmarks.get_Item(ref oEndOfDoc).Range;
                    oPara2 = app.ActiveDocument.Content.Paragraphs.Add(ref oRng);
                    oPara2.Range.Text = copyText;
                    oPara2.Range.InsertParagraphAfter();
                    Object newFileName = Path.Combine(_fileInfo.DirectoryName, _fileInfo.Name);
                    app.ActiveDocument.SaveAs2(newFileName);
                    app.ActiveDocument.Close();
                }


                return true;
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
            }
            finally
            {
                if (app != null)
                {
                    app.Quit();
                }
            }

            return false;


        }
    }
}
