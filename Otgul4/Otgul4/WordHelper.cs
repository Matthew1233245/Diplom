using System;
using System.Collections.Generic;
using System.Data.OleDb;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Word = Microsoft.Office.Interop.Word;

namespace Otgul4
{
    internal class WordHelper
    {

        private FileInfo _fileInfo;

        public WordHelper(string fileName) 
        {


            if (File.Exists(fileName))
            {
                _fileInfo = new FileInfo(fileName);

            }
            else
            {
                throw new ArgumentException("Файл не обнаружен");
            }
        }

        internal bool Process(Dictionary<string, string> items)
        {
            Word.Application app = null;
            SaveFileDialog saveFileDialog = new SaveFileDialog();
            try
            {
                app = new Word.Application();
                Object file = _fileInfo.FullName;

                Object missing = Type.Missing;

                app.Documents.Open(file);

                foreach (var item in items)
                {
                    Word.Find find = app.Selection.Find;
                    find.Text = item.Key;
                    find.Replacement.Text = item.Value;

                    Object wrap = Word.WdFindWrap.wdFindContinue;
                    Object replace = Word.WdReplace.wdReplaceAll;

                    find.Execute(FindText: Type.Missing,
                        MatchCase: false,
                        MatchWholeWord: false,
                        MatchWildcards: false,
                        MatchSoundsLike: missing,
                        MatchAllWordForms: false,
                        Forward: true,
                        Wrap: wrap,
                        Format: false,
                        ReplaceWith: missing, Replace: replace);
                }

                // Настройка SaveFileDialog
                saveFileDialog.Filter = "Word Documents (*.docx)|*.docx|All files (*.*)|*.*";
                saveFileDialog.FileName = DateTime.Now.ToString("ddMMyyyy HHmmss") + _fileInfo.Name;
                saveFileDialog.InitialDirectory = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments);

                // Показ диалога сохранения файла
                if (saveFileDialog.ShowDialog() == DialogResult.OK)
                {
                    Object newFileName = saveFileDialog.FileName;
                    app.ActiveDocument.SaveAs2(newFileName);
                    app.ActiveDocument.Close();

                    /*// Открываем созданный файл
                    app.Documents.Open(newFileName);
                    app.Visible = true; // Делаем приложение Word видимым*/

                    // Открываем сохраненный файл
                    System.Diagnostics.Process.Start(saveFileDialog.FileName);

                    return true;
                }
                else
                {
                    // Если пользователь отменил сохранение
                    app.ActiveDocument.Close(Word.WdSaveOptions.wdDoNotSaveChanges);
                    return false;
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
            }
            finally
            {
                // Убедитесь, что приложение Word закрыто, даже если произошла ошибка
                if (app != null)
                {
                    app.Quit();
                }
            }

            return false;

        }
    }
}
