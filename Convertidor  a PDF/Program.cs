using System;
using System.IO;
using System.Linq;
using System.Diagnostics;
using System.Threading.Tasks;
using HtmlAgilityPack;
using iText.Kernel.Pdf;
using PuppeteerSharp;
using PuppeteerSharp.Media;
using iText.Kernel.Pdf;
using iText.Kernel.Pdf.Annot;
using iText.Kernel.Pdf.Action;
using iText.Kernel;
using iText.Layout.Element;

class Program
{
    static async Task Main(string[] args)
    {
        // Si no se pasan args, usar valores de prueba
        if (args.Length < 3)
        {
            args = new string[]
            {
                @"C:\Exportaciones\Expediente_465674-2025-RI",       // Carpeta base
                "Expediente.html",           // HTML original
                "ExpedienteFinal.pdf"        // PDF final
            };

            //Console.WriteLine("❌ Error: Se requieren 3 parámetros:");
            //Console.WriteLine("   1. Carpeta base del expediente");
            //Console.WriteLine("   2. Nombre del archivo HTML");
            //Console.WriteLine("   3. Nombre del archivo PDF de salida");
            //Console.WriteLine();
            //Console.WriteLine(@"Ejemplo:");
            //Console.WriteLine(@"   ConvertidorPDF.exe ""C:\Exportaciones\Expediente_437994-2024-RI"" ""Expediente.html"" ""ExpedienteFinal.pdf""");
            //return;
        }

        string baseFolder = args[0];
        string htmlOriginal = Path.Combine(baseFolder, args[1]);
        string htmlModificado = Path.Combine(baseFolder, "Expediente_modificado.html");
        string pdfFinal = Path.Combine(baseFolder, args[2]);

        Console.WriteLine("📄 Leyendo HTML y agregando número de páginas...");
        AgregarCantidadPaginasAPDFs(htmlOriginal, baseFolder, htmlModificado);

        Console.WriteLine("🖨 Convirtiendo HTML a PDF...");
        await ConvertHtmlToPdf(htmlModificado, pdfFinal);

        // 🔧 Limpiar URIs absolutas convertidas por Puppeteer
        Console.WriteLine("🔗 Corrigiendo enlaces para que sean relativos...");
        //ReemplazarURIsEnPDF(pdfFinal, "relative://", "file:");

        LimpiarURIsRelativasConLaunch2(pdfFinal);
        // LimpiarURIsAbsolutas(pdfFinal);

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

        //var pdfLinks = doc.DocumentNode.SelectNodes("//a[contains(@href, '.pdf')]");
        var pdfLinks = doc.DocumentNode.SelectNodes("//a[contains(@href, '.')]");
        if (pdfLinks == null)
        {
            Console.WriteLine("⚠ No se encontraron enlaces a PDF.");
            doc.Save(outputPath); // 🛠 Guarda el HTML igual
            Console.WriteLine($"✅ HTML sin modificaciones guardado en: {outputPath}");
            return;
        }

        foreach (var link in pdfLinks)
        {
            string hrefRelativo = link.GetAttributeValue("href", "").Replace("\\", "/"); // normalizamos slashes

            string rutaPdf = Path.Combine(baseFolder, hrefRelativo);

            if (File.Exists(rutaPdf))
            {
                string ext = Path.GetExtension(hrefRelativo).ToLower();
                Console.WriteLine(@"encuentra link para comvertir");
                try
                {
                    if (ext == ".pdf")
                    {
                        using var reader = new PdfReader(rutaPdf);
                        using var pdfDoc = new PdfDocument(reader);
                        int paginas = pdfDoc.GetNumberOfPages();

                        string textoActual = link.InnerText;
                        link.InnerHtml = $"{paginas} páginas - {textoActual}";
                    }

                    // forzar file: + ruta relativa
                    Console.WriteLine(@"Reemplaza el link  por File");
                    //link.SetAttributeValue("href", $"file:{hrefRelativo}");
                    link.SetAttributeValue("href", $"relative:{hrefRelativo}"); /// con LimpiarURIsRelativasConLaunch
                    //link.SetAttributeValue("href", $"relative://{hrefRelativo}");
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



    static PdfAction CrearAccionSmart(string path)
    {
        string ext = Path.GetExtension(path).ToLower();

        if (ext == ".pdf")
        {
            // GoToR (abre en nueva pestaña si es PDF)
            return PdfAction.CreateGoToR(path, 1, true);
        }
        else
        {
            // LaunchAction para cualquier otro tipo
            var launchAction = new PdfAction();
            var launchDict = launchAction.GetPdfObject();
            launchDict.Put(PdfName.S, PdfName.Launch);
            launchDict.Put(PdfName.F, new PdfString(path));
            return launchAction;
        }
    }

    static void LimpiarURIsRelativasConLaunch2(string pdfPath)
    {
        string tempPath = pdfPath + ".tmp";

        using var reader = new PdfReader(pdfPath);
        using var writer = new PdfWriter(tempPath);
        using var pdfDoc = new PdfDocument(reader, writer);

        for (int i = 1; i <= pdfDoc.GetNumberOfPages(); i++)
        {
            var page = pdfDoc.GetPage(i);
            var annots = page.GetAnnotations();

            foreach (var annot in annots)
            {
                if (annot is PdfLinkAnnotation linkAnnot)
                {
                    var dict = linkAnnot.GetPdfObject();
                    var oldAction = dict.GetAsDictionary(PdfName.A);

                    if (oldAction != null && oldAction.Get(PdfName.S)?.ToString() == "/URI")
                    {
                        string uri = oldAction.GetAsString(PdfName.URI)?.ToString() ?? "";

                        if (uri.StartsWith("relative:"))
                        {
                            string relativePath = uri.Replace("relative:", "").Replace("\\", "/");

                            // 💡 Usar lógica inteligente según tipo de archivo
                            var action = CrearAccionSmart(relativePath);
                            linkAnnot.SetAction(action);

                            Console.WriteLine($"🔗 Acción asignada a: {relativePath}");
                        }
                    }
                }
            }
        }

        // JavaScript de advertencia
        var js = @"
            if (app.viewerType !== 'Reader' && app.viewerType !== 'Exchange') {
                app.alert(' Este PDF se ve mejor en Adobe Acrobat.\nAbra este archivo con Acrobat Reader para una mejor experiencia.');
            }";
        pdfDoc.GetCatalog().SetOpenAction(PdfAction.CreateJavaScript(js));

        pdfDoc.Close();
        File.Delete(pdfPath);
        File.Move(tempPath, pdfPath);
    }
}
