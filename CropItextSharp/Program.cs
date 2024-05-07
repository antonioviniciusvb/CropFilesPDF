using iTextSharp.text;
using iTextSharp.text.pdf;
using iTextSharp.text.pdf.parser;
using System;
using System.IO;
using System.Text;
using System.Text.RegularExpressions;

namespace PDFCropExample
{
    class Program
    {
        //Arquivo de Entrada
        public static string inputFile = @"input.pdf";
        public static string tmpFolder = @"d:\tmp";
        public static string processFolder = @"d:\process";

        //Dimensões para as páginas Impares
        public static float[] rectanguleOdd =
            {
                Utilities.MillimetersToPoints(96),
                Utilities.MillimetersToPoints(96),
                Utilities.MillimetersToPoints(98)
            };

        //Dimensões para as páginas Pares
        public static float[] rectanguleEven =
            {
                Utilities.MillimetersToPoints(101),
                Utilities.MillimetersToPoints(98),
                Utilities.MillimetersToPoints(98)
            };

        //Largura da página A4
        public static float widthDocument = Utilities.MillimetersToPoints(210);

        //Dimensões do retangulo de corte
        public static float[] x = { 0, 0, 0 };
        public static float[] y = { 0, 0, 0 };
        public static float[] right = { widthDocument, widthDocument, widthDocument };
        public static float[] top = { 0, 0, 0 };

        static void Main(string[] args)
        {
            ClearCreateFolder(tmpFolder);

            ClearCreateFolder(processFolder);

            Split();

            StringBuilder keyFiles = GetKeysFilename();

            MoveFiles(keyFiles);
            ExportReport(keyFiles);

            Console.WriteLine("");
            Console.WriteLine("");
            Console.WriteLine("**** Processamento finalizado com sucesso!");
            Console.WriteLine("");
            Console.WriteLine($"Verifique a pasta: {processFolder}");
            Console.ReadKey();
        }

        public static void ExportReport(StringBuilder keyFiles)
        {
            var pdfProcess = $"{processFolder}\\Process-FileNamesPdf.txt";

            File.WriteAllText(pdfProcess, $"{keyFiles}");
            Console.WriteLine($"Passo 4 - File: [{pdfProcess}] - Exportado");
        }

        public static void MoveFiles(StringBuilder keyFiles)
        {
            string[] lines = Regex.Split($"{keyFiles}".Trim(), "\n");

            foreach (var line in lines)
            {
                string[] fields = Regex.Split($"{line}", ";");
                File.Move(fields[0], fields[1]);
                Console.WriteLine($"Passo 3 - File: [{fields[1]}] - Renomeado");
            }
        }

        public static StringBuilder GetKeysFilename()
        {
            var numberPages = GetNumberPages(inputFile) * 3;
            var lines = new StringBuilder();
            var contribuinte = new StringBuilder();
            var inscricao = new StringBuilder();
            var files = new StringBuilder();
            int counter = 1;

            for (int i = 1, corte = 0, lamina = 1; i <= numberPages; i++, corte++, lamina++)
            {
                var tmpFile = GetTemporyFileName(i);

                if (corte > 2)
                    corte = 0;

                if (lamina > 6)
                {
                    lamina = 1;
                    counter++;
                }


                Rectangle rect = new Rectangle(x[corte], y[corte], right[corte], top[corte]);
                RenderFilter[] filter = { new RegionTextRenderFilter(rect) };

                ITextExtractionStrategy strategy = new FilteredTextRenderListener(new LocationTextExtractionStrategy(), filter);

                using (PdfReader reader = new PdfReader(tmpFile))
                {
                    lines.Clear();
                    lines.AppendLine(tmpFile);
                    lines.Append(PdfTextExtractor.GetTextFromPage(reader, 1, strategy));

                    if (lamina == 1)
                    {
                        var findContribuinte = FindInText(lines, @"^contribuinte: (?<name>[^\n]+)",
                            RegexOptions.IgnoreCase | RegexOptions.Multiline);

                        if (findContribuinte.Count == 1)
                        {
                            contribuinte.Clear();
                            contribuinte.Append(findContribuinte[0].Groups["name"].Value);
                        }

                        var findInscricao = FindInText(lines, @"^(?:Matricula:[^\n]+\n(?<insc>\w{1,5}[.]\w{1,5}[.]\w{1,5}[.]\w{1,5}[.]\w{1,5}))",
                                           RegexOptions.IgnoreCase | RegexOptions.Multiline);

                        if (findInscricao.Count == 1)
                        {
                            inscricao.Clear();
                            inscricao.Append(findInscricao[0].Groups["insc"].Value);
                        }
                    }
                    else
                    {
                        //var findContribuinte = FindInText(lines, $"{contribuinte}",
                        //                   RegexOptions.IgnoreCase | RegexOptions.Multiline).Count > 0;
                        var findContribuinte = true;

                        var findInscricao = FindInText(lines, $"^(?:Inscrição|Matricula): {inscricao}",
                                           RegexOptions.IgnoreCase | RegexOptions.Multiline).Count > 0;

                        if (!(findContribuinte && findContribuinte))
                            throw new Exception("Chaves não localizadas");
                    }

                    Console.WriteLine($"Passo 2 - Registro: [{counter}-{lamina}] - Validado");

                    files.Append($"{tmpFile};{processFolder}\\Process-{inscricao}_{counter}-{lamina}.pdf\n");
                }
            }

            return files;
        }

        private static MatchCollection FindInText(StringBuilder lines, string pattern, RegexOptions options)
        {
            var find = Regex.Matches(lines.ToString(), pattern, options);

            if (find.Count == 0)
            {
                //Encontrar se existe uma linha que começa com Matricula
                //Se existir pegar a linha toda e cria um grupo de captura
                var findInscricao = Regex.Matches(lines.ToString(), @"^(?:Matricula:)(?<insc>[^\n]+\n)", options);

                //Remove os espaços em brancos
                var auxInsc = $"Matricula: {findInscricao[0].Groups["insc"].Value.Replace(" ", "")}";

                //Tenta novamente encontrar o pattern
                find = Regex.Matches(auxInsc, pattern, options);

                if (find.Count == 0)
                    throw new Exception("Pattern não foi encontrado");
            }

            return find;
        }

        public static void Split()
        {
            // Abre o arquivo original
            using (PdfReader reader = new PdfReader(inputFile))
            {
                // Percorre todas as páginas
                for (int i = 1, pages = 1; i <= reader.NumberOfPages; i++)
                {
                    //Iniciando os recortes
                    for (int j = 0; j < 3; j++)
                    {
                        Console.WriteLine($"Passo 1 - File Temp: {pages}");
                       
                        var tmpFile = GetTemporyFileName(pages++);

                        using (FileStream fs = new FileStream(tmpFile, FileMode.Create))
                        {
                            Document document = new Document();

                            using (PdfWriter writer = PdfWriter.GetInstance(document, fs))
                            {
                                document.Open();

                                // Obtém a página original
                                PdfImportedPage page = writer.GetImportedPage(reader, i);

                                //Define o tamanho do recorte
                                GetRectangleHeight(i, page);

                                Rectangle rectangle = new Rectangle(x[j], y[j], right[j], top[j]);

                                document.SetPageSize(rectangle);

                                document.NewPage();

                                PdfContentByte cb = writer.DirectContent;

                                //Adiciona o conteudo na página
                                cb.AddTemplate(page, 0, 0);

                                document.Close();
                            }
                        }
                    }

                }
            }
        }

        public static string GetTemporyFileName(int index)
        {
            return $"{tmpFolder}\\Tmp_{index}.pdf";
        }

        public static int GetNumberPages(string inputFile)
        {
            int numberPages = 0;

            using (PdfReader reader = new PdfReader(inputFile))
            {
                numberPages = reader.NumberOfPages;
            }

            return numberPages;
        }

        public static void ClearCreateFolder(string folder)
        {
            if (Directory.Exists(folder))
                Directory.Delete(folder, true);

            Directory.CreateDirectory(folder);
        }

        public static void GetRectangleHeight(int numberPage, PdfImportedPage page)
        {

            if (numberPage % 2 == 0)
            {
                y[0] = page.Height;
                y[1] = page.Height - rectanguleEven[0];
                y[2] = y[1] - rectanguleEven[1];

                top[0] = y[0] - rectanguleEven[0];
                top[1] = top[0] - rectanguleEven[1];
                top[2] = top[1] - rectanguleEven[2];
            }
            else
            {

                y[0] = page.Height;
                y[1] = page.Height - rectanguleOdd[0];
                y[2] = y[1] - rectanguleOdd[1];

                top[0] = y[0] - rectanguleOdd[0];
                top[1] = top[0] - rectanguleOdd[1];
                top[2] = top[1] - rectanguleOdd[2];
            }
        }
    }
}
