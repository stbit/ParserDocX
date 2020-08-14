using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.Xml.Linq;
using Xceed.Document.NET;
using Xceed.Words.NET;

namespace ParserDocx
{
    public class ConvertDocx
    {
        public static ConvertDocx Parse(string filePath)
        {
            return new ConvertDocx(filePath);
        }

        private string FilePath { get; set; }

        /// <summary>
        /// Рабочая программа
        /// </summary>
        private DocX WorkProgramms { get; set; }

        /// <summary>
        /// Аннотация РП
        /// </summary>
        private DocX Annotations { get; set; }

        /// <summary>
        /// ФОС
        /// </summary>
        private DocX Fos { get; set; }

        public ConvertDocx(string filePath)
        {
            FilePath = filePath;
            WorkProgramms = DocX.Load(filePath);
            Annotations = DocX.Load(filePath);
            Fos = DocX.Load(filePath);

            ParseWorkProgramms();
            ParseAnnotations();
            ParseFos();
        }

        /// <summary>
        /// Из xml вырезаются все теги, переносы строк и пробелы (весь текст без пробелов в одну строку)
        /// </summary>
        /// <param name="node"></param>
        /// <returns></returns>
        private string NodeToString(XNode node)
        {
            return Regex.Replace(node.ToString(), @"<.*?>|\s+", String.Empty).ToLower();
        }

        /// <summary>
        /// Добавляем титульный лист
        /// </summary>
        /// <param name="document"></param>
        private void AddTitle(DocX document)
        {
            document.Xml.Elements().First().FirstNode.Remove();
            
            var title = document.InsertParagraph("Титульный лист");

            title.FontSize(18).Bold().Alignment = Alignment.center;

            // Выводим в xml на верхний уровень для корректной работы InsertPageBreakAfterSelf 
            document.RemoveParagraph(title);
            document.Xml.AddFirst(title.Xml);
            
            title.InsertPageBreakAfterSelf();
        }

        /// <summary>
        /// Удаляем первые 3 страницы до содержания
        /// </summary>
        /// <param name="document"></param>
        private void TruncateBegin(DocX document)
        {
            document.Xml.Elements()
                .TakeWhile(node => !NodeToString(node).Contains("содержание(рабочаяпрограмма)"))
                .ToList()
                .ForEach(node => node.Remove());
        }

        /// <summary>
        /// Преобразования xml для "Рабочая программа"
        /// </summary>
        private void ParseWorkProgramms()
        {
            TruncateBegin(WorkProgramms);

            // Удаляем ФОС страницу (стр 5)
            WorkProgramms.Xml.Elements()
                .Skip(2) // Пропускаем контент страницы и описание страницы (w:tlb и w:p) - страница 4
                .TakeWhile(node => !NodeToString(node).Contains("местодисциплинывструктуре"))
                .ToList()
                .ForEach(node => node.Remove());

            AddTitle(WorkProgramms);
        }

        /// <summary>
        /// Преобразования xml для "Аннотация РП"
        /// </summary>
        private void ParseAnnotations()
        {
            TruncateBegin(Annotations);

            // Удаляем стр 4,5
            Annotations.Xml.Elements()
                .TakeWhile((node, index) => index == 0 || !NodeToString(node).Contains("местодисциплинывструктуре"))
                .ToList()
                .ForEach(node => node.Remove());

            // Разворачиваем xml node (с 6. Фонд оценочных средств по дисциплине) и удаляем лишнее
            Annotations.Xml.Elements()
                .Where(node => NodeToString(node).Contains("фондоценочныхсредствподисциплине"))
                .SelectMany(node => node.Elements())
                .SkipWhile(node => !NodeToString(node).Contains("фондоценочныхсредствподисциплине"))
                .ToList()
                .ForEach(node => node.Remove());

            AddTitle(Annotations);
        }

        /// <summary>
        /// Преобразования xml для "ФОС"
        /// </summary>
        private void ParseFos()
        {
            TruncateBegin(Fos);

            // Удаляем 4 страницу
            Fos.Xml.Elements()
                .Take(2)
                .ToList()
                .ForEach(node => node.Remove());

            // Удаляем все до 6 пункта
            Fos.Xml.Elements()
                .Skip(2) // Пропускаем оглавление
                .TakeWhile(node => !NodeToString(node).Contains("фондоценочныхсредствподисциплине"))
                .Concat(
                    Fos.Xml.Elements()
                        .Where(node => NodeToString(node).Contains("фондоценочныхсредствподисциплине"))
                        .SelectMany(node => node.Elements())
                        .TakeWhile(node => !NodeToString(node).Contains("фондоценочныхсредствподисциплине"))
                )
                .ToList()
                .ForEach(node => {
                    node.Remove();
                });

            AddTitle(Fos);
        }


        /// <summary>
        /// Сохранение 3х файлов после преобразований
        /// </summary>
        public void Save()
        {
            string extension = Path.GetExtension(FilePath);
            string documentsDirectory = Path.Combine(Path.GetDirectoryName(FilePath), "dist", Path.GetFileNameWithoutExtension(FilePath));

            if (Directory.Exists(documentsDirectory)) Directory.Delete(documentsDirectory, true);

            Directory.CreateDirectory(documentsDirectory);

            WorkProgramms.SaveAs(Path.Combine(documentsDirectory, "Рабочая программа" + extension));
            Annotations.SaveAs(Path.Combine(documentsDirectory, "Аннотация РП" + extension));
            Fos.SaveAs(Path.Combine(documentsDirectory, "ФОС" + extension));
        }
    }
}
