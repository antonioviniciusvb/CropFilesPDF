using iTextSharp.text;
using iTextSharp.text.pdf;
using System;
using System.Diagnostics;
using System.IO;

namespace PDFCropExample
{
    class Program
    {
        //Arquivo de Entrada
        public static string inputFile = @"input.pdf";

        //Arquivo de Saida
        public static string outputFile = @"output.pdf";

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
        public static float[] top =
            {
                Utilities.MillimetersToPoints(201),
                Utilities.MillimetersToPoints(105),
                0
            };

        static void Main(string[] args)
        {
            // Abre o arquivo original
            using (PdfReader reader = new PdfReader(inputFile))
            {
                Document document = new Document();

                using (FileStream fs = new FileStream(outputFile, FileMode.Create))
                {
                    using (PdfWriter writer = PdfWriter.GetInstance(document, fs))
                    {
                        document.Open();

                        // Percorre todas as páginas
                        for (int i = 1; i <= reader.NumberOfPages; i++)
                        {
                            // Obtém a página original
                            PdfImportedPage page = writer.GetImportedPage(reader, i);

                            //Define o tamanho do recorte
                            GetRectangleHeight(i, page);


                            //Iniciando os recortes
                            for (int j = 0; j < 3; j++)
                            {
                                Rectangle rectangle = new Rectangle(x[j], y[j], right[j], top[j]);

                                document.SetPageSize(rectangle);

                                document.NewPage();

                                PdfContentByte cb = writer.DirectContent;

                                //Adiciona o conteudo na página
                                cb.AddTemplate(page, 0, 0);
                            }
                        }

                        document.Close();
                    }
                }
            }

            //Abre o documento no Reader Padrão
            Process.Start(outputFile);

            Console.WriteLine("Recorte concluído. Novo arquivo criado: " + outputFile);
        }

        private static void GetRectangleHeight(int numberPage, PdfImportedPage page)
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
