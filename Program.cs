using System;
using System.IO;
using System.Collections.Generic;
using System.Linq;
using iText.Kernel.Pdf;
using iText.Layout;
using iText.Layout.Element;
using iText.Kernel.Pdf.Canvas.Parser;
using System.Reflection.PortableExecutable;
using System.Text.RegularExpressions;

class Program
{
    static void Main()
    {
        string sourceFile = @"C:\Users\Lucas Souza\Downloads\Relatórios Elétrica - Abril 2024\mec";
        string destinationFolder = @"C:\Users\Lucas Souza\Downloads\Relatórios - Abril 2024\mecanicaprev";

        Console.WriteLine("Verificando caminhos...");
        Console.WriteLine($"Caminho do arquivo fonte: {sourceFile}");
        Console.WriteLine($"Caminho do diretório de destino: {destinationFolder}");

        if (!File.Exists(sourceFile))
        {
            Console.WriteLine("Arquivo fonte não encontrado.");
            return;
        }

        if (!Directory.Exists(destinationFolder))
        {
            Console.WriteLine("Diretório de destino não encontrado.");
            return;
        }

        try
        {
            SplitPdfByOS(sourceFile, destinationFolder);
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Erro ao processar o arquivo: {ex.Message}");
        }
    }

    static void SplitPdfByOS(string sourceFile, string destinationFolder)
    {
        try
        {
            using (PdfReader pdfReader = new PdfReader(sourceFile))
            {
                PdfDocument pdfDoc = new PdfDocument(pdfReader);
                int numberOfPages = pdfDoc.GetNumberOfPages();

                Dictionary<string, List<int>> osPages = new Dictionary<string, List<int>>();
                string currentOS = null;

                for (int i = 1; i <= numberOfPages; i++)
                {
                    var pageContent = PdfTextExtractor.GetTextFromPage(pdfDoc.GetPage(i));

                    Console.WriteLine($"Processando página {i}/{numberOfPages}");

                    // Regex para encontrar número de OS
                    var match = System.Text.RegularExpressions.Regex.Match(pageContent, @"OS: (\d+)");
                    if (match.Success)
                    {
                        currentOS = match.Groups[1].Value;
                        Console.WriteLine($"Encontrado OS: {currentOS} na página {i}");

                        if (!osPages.ContainsKey(currentOS))
                        {
                            osPages[currentOS] = new List<int>();
                        }
                    }

                    if (currentOS != null)
                    {
                        osPages[currentOS].Add(i);
                    }
                }

                foreach (var osEntry in osPages)
                {
                    string osNumber = osEntry.Key;
                    List<int> pages = osEntry.Value;

                    string newFileName = $"Mecânica - Preventiva - OS {osNumber}.pdf";
                    string newFilePath = Path.Combine(destinationFolder, newFileName);

                    try
                    {
                        Console.WriteLine($"Criando arquivo: {newFileName}");

                        using (PdfWriter pdfWriter = new PdfWriter(newFilePath))
                        {
                            using (PdfDocument newPdfDoc = new PdfDocument(pdfWriter))
                            {
                                foreach (int pageNum in pages)
                                {
                                    pdfDoc.CopyPagesTo(pageNum, pageNum, newPdfDoc);
                                }
                            }
                        }

                        Console.WriteLine($"Arquivo criado: {newFilePath}");
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine($"Erro ao criar arquivo {newFileName}: {ex.Message}");
                    }
                }
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Erro ao processar o PDF: {ex.Message}");
            throw;
        }

    }
