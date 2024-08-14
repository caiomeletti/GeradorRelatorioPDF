//outra opção
//https://ironpdf.com/blog/compare-to-other-components/itextsharp/

using iTextSharp.text;
using iTextSharp.text.pdf;
using System.Diagnostics;

namespace GeradorRelatorioPDF
{
    public class Program
    {
        static List<Pessoa> pessoas = new();

        static void Main(string[] args)
        {
            pessoas = DesserializarPessoas();
            //foreach (var pessoa in pessoas)
            //{
            //    Console.WriteLine($"{pessoa.IdPessoa} - {pessoa.Nome} {pessoa.Sobrenome}");
            //}
            GerarRelatorioPDF(100);
        }

        static void GerarRelatorioPDF(int qtdePessoas)
        {
            var pessoasSelecionadas = pessoas.Take(qtdePessoas).ToList();
            if (pessoasSelecionadas.Count > 0)
            {
                //calculo qtde total de paginas
                int totalPaginas = 1;
                int totalLinhas = pessoasSelecionadas.Count;
                if (totalLinhas > 24)
                {
                    totalPaginas += (int)Math.Ceiling((totalLinhas - 24) / 29F);
                }

                //configuracao do documento pdf
                var pxPorMm = 72 / 25.2F;
                var pdf = new Document(PageSize.A4, 15 * pxPorMm, 15 * pxPorMm, 15 * pxPorMm, 20 * pxPorMm);
                var nomeArquivo = $"pessoas.{DateTime.Now:yyyy.MM.dd.hh.mm.ss}.pdf";
                var arquivo = new FileStream(nomeArquivo, FileMode.Create);
                var writer = PdfWriter.GetInstance(pdf, arquivo);
                writer.PageEvent = new EventosDePagina(totalPaginas);
                pdf.Open();

                var fontBase = BaseFont.CreateFont(BaseFont.HELVETICA, BaseFont.CP1252, false);
                //titulo
                var fontParagrafo = new Font(fontBase, 32, Font.NORMAL, BaseColor.Black);
                var titulo = new Paragraph("Relatório de pessoas\n\n", fontParagrafo)
                {
                    Alignment = Element.ALIGN_LEFT,
                    SpacingAfter = 4
                };
                pdf.Add(titulo);

                //imagem no titulo
                var caminhoImagem = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "img\\youtube.png");
                if (File.Exists(caminhoImagem))
                {
                    Image logo = Image.GetInstance(caminhoImagem);
                    float razaoAlturaLargura = logo.Width / logo.Height;
                    float alturaLogo = 32;
                    float larguraLogo = alturaLogo * razaoAlturaLargura;
                    logo.ScaleToFit(larguraLogo, alturaLogo);
                    var margemEsquerda = pdf.PageSize.Width - pdf.RightMargin - larguraLogo;
                    var margemTopo = pdf.PageSize.Height - pdf.TopMargin - 54;
                    logo.SetAbsolutePosition(margemEsquerda, margemTopo);
                    writer.DirectContent.AddImage(logo, false);
                }

                //tabela de dados
                var tabela = new PdfPTable(5);
                float[] larguraColunas = { 0.6f, 2f, 1.5f, 1f, 1f };
                tabela.SetWidths(larguraColunas);
                tabela.DefaultCell.BorderWidth = 0;
                tabela.WidthPercentage = 100;

                var currentRow = 0;
                //cabecalho da tabela
                tabela.AddCell(CriarCelulaTexto(fontBase, "Código", currentRow, PdfCell.ALIGN_CENTER, negrito: true));
                tabela.AddCell(CriarCelulaTexto(fontBase, "Nome", currentRow, PdfCell.ALIGN_CENTER, negrito: true));
                tabela.AddCell(CriarCelulaTexto(fontBase, "Profissão", currentRow, PdfCell.ALIGN_CENTER, negrito: true));
                tabela.AddCell(CriarCelulaTexto(fontBase, "Salário", currentRow, PdfCell.ALIGN_CENTER, negrito: true));
                tabela.AddCell(CriarCelulaTexto(fontBase, "Empregado", currentRow, PdfCell.ALIGN_CENTER, negrito: true));

                //dados da tabela
                foreach (var pessoa in pessoas)
                {
                    currentRow++;
                    tabela.AddCell(CriarCelulaTexto(fontBase, pessoa.IdPessoa.ToString("D6"), currentRow, PdfCell.ALIGN_CENTER));
                    tabela.AddCell(CriarCelulaTexto(fontBase, pessoa.Nome + " " + pessoa.Sobrenome, currentRow));
                    tabela.AddCell(CriarCelulaTexto(fontBase, pessoa.Profissao.Nome, currentRow, PdfCell.ALIGN_CENTER));
                    tabela.AddCell(CriarCelulaTexto(fontBase, pessoa.Salario.ToString("C2"), currentRow, PdfCell.ALIGN_RIGHT));
                    //tabela.AddCell(CriarCelulaTexto(fontBase, pessoa.Empregado ? "Sim" : "Não", currentRow, PdfCell.ALIGN_CENTER));
                    var caminhoImagemCelula = Path.Combine(AppDomain.CurrentDomain.BaseDirectory,
                        pessoa.Empregado ? "img\\happy.png" : "img\\sad.png");
                    tabela.AddCell(CriarCelulaImagem(caminhoImagemCelula, currentRow, 20, 20));
                }
                pdf.Add(tabela);

                pdf.Close();
                arquivo.Close();

                var caminhoPDF = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, nomeArquivo);
                if (File.Exists(caminhoPDF))
                {
                    Process.Start(new ProcessStartInfo()
                    {
                        Arguments = $"/c start {caminhoPDF}",
                        FileName = "cmd.exe",
                        CreateNoWindow = true
                    });
                }
            }
        }

        private static PdfPCell CriarCelulaImagem(string caminhoImagem,
                                                  int linhaAtual,
                                                  int larguraImagem,
                                                  int alturaImagem,
                                                  int alturaCelula = 25)
        {
            if (File.Exists(caminhoImagem))
            {
                Image imagem = Image.GetInstance(caminhoImagem);
                imagem.ScaleToFit(larguraImagem, alturaImagem);

                var celula = new PdfPCell(imagem)
                {
                    HorizontalAlignment = PdfPCell.ALIGN_CENTER,
                    VerticalAlignment = PdfPCell.ALIGN_MIDDLE,
                    Border = 0,
                    BorderWidthBottom = 1,
                    FixedHeight = alturaCelula,
                    BackgroundColor = linhaAtual % 2 == 0
                        ? BaseColor.White
                        : new BaseColor(0.95f, 0.95f, 0.95f)
                };
                return celula;
            }
            else
            {
                return new PdfPCell();
            }
        }

        private static PdfPCell CriarCelulaTexto(BaseFont fontBase,
                                                 string texto,
                                                 int linhaAtual = 0,
                                                 int alinhamentoHz = PdfPCell.ALIGN_LEFT,
                                                 bool negrito = false,
                                                 bool italico = false,
                                                 int tamanhoFonte = 12,
                                                 int alturaCelula = 25)
        {
            int estilo = Font.NORMAL;
            if (negrito && italico)
                estilo = Font.BOLDITALIC;
            else if (negrito)
                estilo = Font.BOLD;
            else if (italico)
                estilo = Font.ITALIC;

            var fonteCelula = new Font(fontBase, tamanhoFonte, estilo, BaseColor.Black);

            var celula = new PdfPCell(new Phrase(texto, fonteCelula))
            {
                HorizontalAlignment = alinhamentoHz,
                VerticalAlignment = PdfPCell.ALIGN_MIDDLE,
                Border = 0,
                BorderWidthBottom = 1,
                FixedHeight = alturaCelula,
                PaddingBottom = 5,
                BackgroundColor = linhaAtual % 2 == 0
                    ? BaseColor.White
                    : new BaseColor(0.95f, 0.95f, 0.95f)
            };
            return celula;
        }

        public static List<Pessoa> DesserializarPessoas()
        {
            List<Pessoa> pessoas = new()
            {
                new Pessoa
                {
                    IdPessoa = 1,
                    Nome = "Obiwan",
                    Sobrenome = "Kenobi",
                    Profissao = new Profissao
                    {
                        IdProfissao = 1,
                        Nome = "Jedi Master"
                    },
                    Empregado = true,
                    Salario = 1500
                },
                new Pessoa
                {
                    IdPessoa = 2,
                    Nome = "Anakin",
                    Sobrenome = "Skywalker",
                    Profissao = new Profissao
                    {
                        IdProfissao = 2,
                        Nome = "Jedi"
                    },
                    Empregado = true,
                    Salario = 500
                },
                new Pessoa
                {
                    IdPessoa = 3,
                    Nome = "Asoka",
                    Sobrenome = "Tano",
                    Profissao = new Profissao
                    {
                        IdProfissao = 3,
                        Nome = "Padawan"
                    },
                    Empregado = true,
                    Salario = 300
                },
                new Pessoa
                {
                    IdPessoa = 4,
                    Nome = "General",
                    Sobrenome = "Grievous",
                    Profissao = new Profissao
                    {
                        IdProfissao = 4,
                        Nome = "Jedi Hunter"
                    },
                    Empregado = false,
                    Salario = 1300
                },
                new Pessoa
                {
                    IdPessoa = 5,
                    Nome = "Leia",
                    Sobrenome = "Organa",
                    Profissao = new Profissao
                    {
                        IdProfissao = 5,
                        Nome = "Princesa"
                    },
                    Empregado = false,
                    Salario = 1
                },
                new Pessoa
                {
                    IdPessoa = 6,
                    Nome = "Han",
                    Sobrenome = "Solo",
                    Profissao = new Profissao
                    {
                        IdProfissao = 6,
                        Nome = "Contrabandista"
                    },
                    Empregado = true,
                    Salario = 2500
                },
                new Pessoa
                {
                    IdPessoa = 6,
                    Nome = "Chewbacca",
                    Sobrenome = "",
                    Profissao = new Profissao
                    {
                        IdProfissao = 7,
                        Nome = "Piloto"
                    },
                    Empregado = true,
                    Salario = 2100
                },
                new Pessoa
                {
                    IdPessoa = 6,
                    Nome = "RD-D2",
                    Sobrenome = "",
                    Profissao = new Profissao
                    {
                        IdProfissao = 8,
                        Nome = "Droid Manutenção"
                    },
                    Empregado = true,
                    Salario = 0
                },
                new Pessoa
                {
                    IdPessoa = 6,
                    Nome = "C-3PO",
                    Sobrenome = "",
                    Profissao = new Profissao
                    {
                        IdProfissao = 9,
                        Nome = "Droid Protocolar"
                    },
                    Empregado = true,
                    Salario = 0
                }
            };
            return pessoas;
        }
    }
}