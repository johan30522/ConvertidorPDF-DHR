using System;
using System.IO;
using System.Linq;
using System.Diagnostics;
using System.Threading.Tasks;
using HtmlAgilityPack;
using iText.Kernel.Pdf;
using PuppeteerSharp;
using PuppeteerSharp.Media;

class Program
{
    static async Task Main(string[] args)
    {
        // Si no se pasan args, usar valores de prueba
        if (args.Length < 3)
        {
            //args = new string[]
            //{
            //    @"C:\Exportaciones\Expediente_437994-2024-RI",       // Carpeta base
            //    "Expediente.html",           // HTML original
            //    "ExpedienteFinal.pdf"        // PDF final
            //};

            Console.WriteLine("❌ Error: Se requieren 3 parámetros:");
            Console.WriteLine("   1. Carpeta base del expediente");
            Console.WriteLine("   2. Nombre del archivo HTML");
            Console.WriteLine("   3. Nombre del archivo PDF de salida");
            Console.WriteLine();
            Console.WriteLine(@"Ejemplo:");
            Console.WriteLine(@"   ConvertidorPDF.exe ""C:\Exportaciones\Expediente_437994-2024-RI"" ""Expediente.html"" ""ExpedienteFinal.pdf""");
            return;
        }

        string baseFolder = args[0];
        string htmlOriginal = Path.Combine(baseFolder, args[1]);
        string htmlModificado = Path.Combine(baseFolder, "Expediente_modificado.html");
        string pdfFinal = Path.Combine(baseFolder, args[2]);

        Console.WriteLine("📄 Leyendo HTML y agregando número de páginas...");
        AgregarCantidadPaginasAPDFs(htmlOriginal, baseFolder, htmlModificado);

        Console.WriteLine("🖨 Convirtiendo HTML a PDF...");
        await ConvertHtmlToPdf(htmlModificado, pdfFinal);

        Console.WriteLine("✅ Proceso finalizado. PDF generado en:");
        Console.WriteLine(pdfFinal);

        // Abrir el archivo automáticamente
        //Process.Start(new ProcessStartInfo(pdfFinal) { UseShellExecute = true });
        var startInfo = new ProcessStartInfo
        {
            FileName = pdfFinal,
            UseShellExecute = true,
            Verb = "open" // ✅ fuerza que use el programa asociado
        };

        Process.Start(startInfo);
    }

    // 🔹 Modificar los enlaces del HTML para incluir la cantidad de páginas
    static void AgregarCantidadPaginasAPDFs(string htmlPath, string baseFolder, string outputPath)
    {
        var doc = new HtmlDocument();
        doc.Load(htmlPath);

        var pdfLinks = doc.DocumentNode.SelectNodes("//a[contains(@href, '.pdf')]");
        if (pdfLinks == null)
        {
            Console.WriteLine("⚠ No se encontraron enlaces a PDF.");
            doc.Save(outputPath); // 🛠 Guarda el HTML igual
            Console.WriteLine($"✅ HTML sin modificaciones guardado en: {outputPath}");
            return;
        }

        foreach (var link in pdfLinks)
        {
            string href = link.GetAttributeValue("href", "").Replace("/", "\\");
            string rutaPdf = Path.Combine(baseFolder, href);

            if (File.Exists(rutaPdf))
            {
                try
                {
                    using var reader = new PdfReader(rutaPdf);
                    using var pdfDoc = new PdfDocument(reader);
                    int paginas = pdfDoc.GetNumberOfPages();

                    string textoActual = link.InnerText;
                    link.InnerHtml = $"{paginas} páginas - {textoActual}";
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"⚠ Error al leer PDF '{rutaPdf}': {ex.Message}");
                }
            }
            else
            {
                Console.WriteLine($"⚠ PDF no encontrado: {rutaPdf}");
            }
        }

        doc.Save(outputPath);
        Console.WriteLine($"✅ HTML actualizado guardado en: {outputPath}");
    }

    // 🔹 Convertir HTML a PDF usando Puppeteer
    static async Task ConvertHtmlToPdf(string htmlPath, string pdfOutput)
    {
        await new BrowserFetcher().DownloadAsync();
        await using var browser = await Puppeteer.LaunchAsync(new LaunchOptions { Headless = true });
        await using var page = await browser.NewPageAsync();

        // ✅ Leer archivo HTML directamente desde el sistema de archivos con ruta absoluta
        string absolutePath = Path.GetFullPath(htmlPath).Replace("\\", "/");
        if (!absolutePath.StartsWith("/")) absolutePath = "/" + absolutePath;
        string fileUrl = $"file://{absolutePath}";

        await page.GoToAsync(fileUrl, WaitUntilNavigation.Networkidle0);

        await page.PdfAsync(pdfOutput, new PdfOptions
        {
            Format = PaperFormat.A4,
            PrintBackground = true,
            //PreferCSSPageSize = true
            MarginOptions = new MarginOptions
            {
                Top = "15mm",
                Bottom = "15mm",
                Left = "0mm",
                Right = "0mm"
            }
        });

        Console.WriteLine("✅ PDF generado correctamente con enlaces activos.");
    }
}
