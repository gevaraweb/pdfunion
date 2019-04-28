using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

using PdfSharp;
using PdfSharp.Drawing;
using PdfSharp.Pdf;
using PdfSharp.Pdf.IO;
using System.Text.RegularExpressions;
using Word = Microsoft.Office.Interop.Word;


namespace pdfunion
{

    public partial class Form1 : Form
    {
        Word._Document doc;
        Word._Application app;

        //https://stackoverflow.com/questions/808670/combine-two-or-more-pdfs
        void CopyPages(PdfDocument from, PdfDocument to)
        {
            for (int i = 0; i < from.PageCount; i++)
            {
                to.AddPage(from.Pages[i]);
            }
        }
        public String translit(string s)
        {
            string ret = "";
            string[] rus = {"а","б","в","г","д","е","ё","ж", "з","и","й","к","л","м", "н",
          "о","п","р","с","т","у","ф","х", "ц", "ч", "ш", "щ",   "ъ", "ы","ь",
          "э","ю", "я", " ", "." };
            string[] eng = {"a","b","v","g","d","e","e","zh","z","i","y","k","l","m","n",
          "o","p","r","s","t","u","f","kh","ts","ch","sh","shch",null,"y",null,
          "e","yu","ya", " ", "." };

            for (int j = 0; j < s.Length; j++)
                for (int i = 0; i < rus.Length; i++)
                    if (s.Substring(j, 1).ToLower().Equals(rus[i])) ret += eng[i];

            return ret;
        }
        public Form1()
        {
            InitializeComponent();
        }


        private void button1_Click(object sender, EventArgs e)
        {

            string dir_temp = Application.StartupPath + "\\Temp\\";

                if (!System.IO.Directory.Exists(dir_temp))
                {
                    System.IO.Directory.CreateDirectory(dir_temp);
                }

            string dir_result = Application.StartupPath + "\\Объединенные файлы\\";

            if (!System.IO.Directory.Exists(dir_result))
            {
                System.IO.Directory.CreateDirectory(dir_result);
            }

            // Сборник лежит там же, где и исполняемый файл .exe
            string FileName = Application.StartupPath + "\\сборник.docx";
            app = new Word.Application();
            app.Visible = true;
            doc = app.Documents.Open(FileName, ReadOnly: true);
            // все страницы doc.SaveAs2(newFileName, 
            // FileFormat: Word.WdSaveFormat.wdFormatPDF ); 

            Word.Table myTable = doc.Tables[1]; // если содержание в первой таблице!!! иначе исправить

            int count = myTable.Rows.Count;
            int Num = 1; //count = 5;
            for (int i = 1; i <= count - 1; i++)
            {


                String articleName = myTable.Rows[i].Cells[1].Range.Text.Trim();
                if (articleName.ToUpper().Contains("СЕКЦИЯ")) continue;

                int len = Math.Min(articleName.Length, 48);
                String name = "№" + Num + " " + translit(articleName.Substring(0, len)) + ".pdf";
                String newFileName = dir_temp + name;
                String newFileUnionName = dir_result + name;

                Num++;

                String st_from = myTable.Rows[i].Cells[2].Range.Text;
                String st_to = myTable.Rows[i + 1].Cells[2].Range.Text;

                st_from = Regex.Replace(st_from, "[^0-9]+", string.Empty);
                st_to = Regex.Replace(st_to, "[^0-9]+", string.Empty);

                int from = Convert.ToInt32(st_from);
                int to = Convert.ToInt32(st_to) - 1;

                doc.ExportAsFixedFormat(newFileName,
                     Word.WdExportFormat.wdExportFormatPDF,
                     false,
                     Word.WdExportOptimizeFor.wdExportOptimizeForPrint,
                     Word.WdExportRange.wdExportFromTo,
                     from,
                     to
               );

                using (PdfDocument one = PdfReader.Open("titul.pdf", PdfDocumentOpenMode.Import))
                using (PdfDocument two = PdfReader.Open(newFileName, PdfDocumentOpenMode.Import))
                using (PdfDocument outPdf = new PdfDocument())
                {
                    CopyPages(one, outPdf);
                    CopyPages(two, outPdf);
                    outPdf.Save(newFileUnionName);
                }


            }

            doc.Close(SaveChanges: false);
            app.Quit();
            MessageBox.Show("Готово!");
        }
    }
}
