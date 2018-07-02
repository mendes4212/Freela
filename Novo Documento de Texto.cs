using System;
using System.Collections.Generic;
using System.Configuration;
using System.Linq;
using System.Threading.Tasks;
using System.Web;
using System.Web.Mvc;
using ContabilizarUsers.Controllers.Utils;
using ContabilizarUsers.Models;
using System.Threading;
using System.IO;
using System.Text;
using ContabilizarUsers.Models.DataSource;
using iTextSharp.text;
using System.Net.Mail;

namespace ContabilizarUsers.Controllers
{
    public class HomeController : Controller
    {
        private Bancos bancos = new Bancos();

        private static string email = "nao-responda@optionsreport.net";

        public ActionResult Index()
        {
            return View();
        }

        public ActionResult UsuariosAtivos()
        {
            var empresa = "";
            var listaColaboradores = new List<ColaboradorSituacao>();
            var listaEmpresasQuantidadeColaboradores = new List<EmpresaQuantidadadeColaborador>();
            try
            {
                //var connString = @"data source = tcp:wlhlj9x8em.database.windows.net,1433;initial catalog = {0}; integrated security = False; user id = prissoftware@wlhlj9x8em;password=Pris6040@usuario; multipleactiveresultsets=True;connect timeout = 60; encrypt=True;trustservercertificate=True;application name = EntityFramework";
                var listabds = bancos.GetReportDatabases();
                for (int i = 12; i < listabds.Count; i++)
                {
                    string connStr = ConfigurationManager.ConnectionStrings[listabds[i]].ConnectionString;
                    using (var ctx = new ContabilizadorContext(connStr))
                    {
                        empresa = listabds[i];
                        listaColaboradores = ConstruirListaColaboradores(DateTime.Now, ctx);
                        var EmpresaQuantidade = new EmpresaQuantidadadeColaborador
                        {
                            NomeEmpresa = listabds[i],
                            Quantidade = listaColaboradores.Count
                        };
                        listaEmpresasQuantidadeColaboradores.Add(EmpresaQuantidade);
                        MemoryStream streamCSV = CriarDocumentoCSV(DateTime.Now, listaColaboradores);
                        MemoryStream streamPDF = CriarDocumentoPDF(DateTime.Now, listaColaboradores,ctx);

                        string corpo = "Documento segue em anexo.";

                        EnviaEmail(corpo, streamPDF, streamCSV,ctx);
                    }
                }
                return RedirectToAction("Index");
            }
            catch(Exception ex)
            {
                MailMessage mail = new MailMessage();
                mail.From = new MailAddress(email);
                mail.To.Add("suporte@pris.com.br");
                mail.Subject = "Erro envio de e-mail rotina Verificar Usuários ativos da " + empresa +" "+DateTime.Today.AddMonths(-1).Month + "/" + DateTime.Today.AddMonths(-1).Year;
                mail.Body = "Exceção: " + ex.Message + ". Exceção interna: " + ex.InnerException.Message + ".";
                return RedirectToAction("Index");
            }
        }

        #region Métodos Auxiliares
        
        public List<SelectListItem> PopulaLista()
        {
            List<SelectListItem> lstBancos = new List<SelectListItem>();
            var banco = new Bancos();
            var lista = banco.GetDatabases();
            foreach(var item in lista)
            {
                var temp = item.Replace("OptionsReport","");
                lstBancos.Add(new SelectListItem { Text = temp, Value = item });
            }
            return lstBancos;
        }

        private List<ColaboradorSituacao> ConstruirListaColaboradores(DateTime dataDocumento, ContabilizadorContext context)
        {
            List<ColaboradorSituacao> lstColaboradores = new List<ColaboradorSituacao>();
            List<Employee> lstEmployee = context.Funcionarios.OrderBy(e => e.Name).ToList();

            foreach (Employee emp in lstEmployee)
            {
                ColaboradorSituacao colaborador = new ColaboradorSituacao()
                {
                    EmployeeID = emp.EmployeeID,
                    EmployeeName = emp.Name,
                    Empresa = emp.Empresa,
                    StatusSistema = emp.Status,
                    LoteEmAberto = false,
                    IncluidoNaContagem = false
                };
                Historical hist = context.Historicos.Where(h => h.Contract.FuncionarioFK == emp.EmployeeID).OrderByDescending(h => h.DataAlteracao).FirstOrDefault();
                Contract contract = context.Contratos.FirstOrDefault(c => c.Time != "Encerrado" && c.FuncionarioFK == colaborador.EmployeeID);
                if (emp.Status == "Ativo")
                {
                    colaborador.IncluidoNaContagem = true;
                    if (contract != null)
                    {
                        colaborador.LoteEmAberto = true;
                    }
                }
                else
                {
                    if (contract != null)
                    {
                        colaborador.IncluidoNaContagem = true;
                        colaborador.LoteEmAberto = true;
                    }
                    else
                    {
                        if (hist != null && hist.DataAlteracao.Year == dataDocumento.AddDays(-1).Year)
                        {
                            colaborador.IncluidoNaContagem = true;
                        }
                    }
                }
                if (hist != null)
                {
                    colaborador.UltimaMovimentacao = hist.DataAlteracao;
                }
                else
                {
                    colaborador.UltimaMovimentacao = null;
                }
                lstColaboradores.Add(colaborador);
            }
            return lstColaboradores;
        }
        #region Envio de Email
        private void EnviaEmail(string corpo, MemoryStream streamPDF, MemoryStream streamCSV, ContabilizadorContext context)
        {
            try
            {
                StockPriceDataSource spds = new StockPriceDataSource(context);
                AdministratorDataSource ads = new AdministratorDataSource(context);

                MailMessage mail = new MailMessage();
                mail.From = new MailAddress(email);

                /*if (spds.GetCompanyName().Equals("Cetip")) //Envia os documetos para os administradores da Cetip
                {
                    List<Administrator> admins = ads.Select().ToList();
                    foreach (Administrator adm in admins)
                    {
                        if (Convert.ToBoolean(adm.RecebeEmailExercicioSolicidato))
                            mail.To.Add(adm.Email);
                    }
                }*/

                /*mail.To.Add("henrique.rocha@pris.com.br");
                mail.To.Add("breno.cueto@pris.com.br");
                mail.To.Add("ed.velho@pris.com.br");
                mail.To.Add("marcela.morais@pris.com.br");
                mail.To.Add("guilherme.saulo@pris.com.br");
                mail.To.Add("andressa.morais@pris.com.br");
                /*mail.To.Add("merielen.santos@pris.com.br"); */
                mail.To.Add("vitor.augusto@pris.com.br");

                streamPDF.Position = 0;
                mail.Subject = "Relatório de contagem de usuários referente a " + DateTime.Today.AddMonths(-1).Month + "/" + DateTime.Today.AddMonths(-1).Year + " " + spds.GetCompanyName();
                mail.Attachments.Add((new System.Net.Mail.Attachment(streamPDF, "Relatório de contagem de usuários -" + DateTime.Today.ToShortDateString() + "-" + spds.GetCompanyName() + ".pdf")));
                mail.Attachments.Add((new System.Net.Mail.Attachment(streamCSV, "Relatório de contagem de usuários -" + DateTime.Today.ToShortDateString() + "-" + spds.GetCompanyName() + ".csv")));
                string body = "Fatura referente ao mês " + DateTime.Today.AddMonths(-1).Month + "/" + DateTime.Today.AddMonths(-1).Year +
                " do Options Report " + spds.GetCompanyName() + "\n\n\n\n" + corpo;

                mail.Body = body;

                Email.SendEmail(mail);
            }
            catch (Exception ex)
            {
                MailMessage mail = new MailMessage();
                mail.From = new MailAddress(email);
                mail.To.Add("suporte@pris.com.br");
                mail.Subject = "Erro envio de e-mail rotina Verificar Usuários ativos " + DateTime.Today.AddMonths(-1).Month + "/" + DateTime.Today.AddMonths(-1).Year;
                mail.Body = "Exceção: " + ex.Message + ". Exceção interna: " + ex.InnerException.Message + ".";
            }
        }
        #endregion
        #region Criação de documentos
        private MemoryStream CriarDocumentoCSV(DateTime dataDocumento, List<ColaboradorSituacao> lstColaboradores)
        {

            MemoryStream stream = new MemoryStream();
            StreamWriter writer = new StreamWriter(stream, Encoding.UTF8);    // using UTF-8 encoding by default

            int countColaboradoresIncluidos = lstColaboradores.Where(c => c.IncluidoNaContagem == true).ToList().Count;

            writer.WriteLine("É considerado beneficiário ativo no sistema aquele que cumpre uma ou mais das seguintes condições:; ; ; ; ;");
            writer.WriteLine("a) Possuir login ativo.; ; ; ; ;");
            writer.WriteLine("b) Possuir contratos com status \"Em Aberto\"; ; ; ; ;");
            writer.WriteLine("c) Possuir alguma movimentação (exercício, resgate ou cancelamento) no ano corrente.; ; ; ; ;");
            writer.WriteLine("; ; ; ; ;");
            writer.WriteLine("Período:;" + 01 + "/" + dataDocumento.AddMonths(-1).Month + "/" + dataDocumento.AddMonths(-1).Year + "; a; " + dataDocumento.AddDays(-1).Day + "/" + dataDocumento.AddDays(-1).Month + "/" + dataDocumento.AddDays(-1).Year + "; ;");
            writer.WriteLine("Número de usuários no período:; " + countColaboradoresIncluidos + "; ; ; ;");
            writer.WriteLine(" ; ; ; ; ;");
            writer.WriteLine("Nome; Empresa; Status no sistema; Possui lotes em aberto?; Data da última movimentação; Incluído na contagem ?");
            foreach (ColaboradorSituacao colaborador in lstColaboradores)
            {
                string incluidoContagem = colaborador.IncluidoNaContagem ? "Sim" : "Não";
                string loteAberto = colaborador.LoteEmAberto ? "Sim" : "Não";
                string dataMovimentacao = "-";
                if (colaborador.UltimaMovimentacao != null)
                {
                    dataMovimentacao = colaborador.UltimaMovimentacao.Value.ToShortDateString();
                }
                writer.WriteLine(colaborador.EmployeeName + ";" + colaborador.Empresa + ";" + colaborador.StatusSistema + ";" + loteAberto + ";" + dataMovimentacao + ";" + incluidoContagem);
            }

            writer.Flush();
            stream.Position = 0;     // read from the start of what was written
            return stream;
        }

        private MemoryStream CriarDocumentoPDF(DateTime dataDocumento, List<ColaboradorSituacao> lstColaboradores,ContabilizadorContext context)
        {

            int countColaboradoresIncluidos = lstColaboradores.Where(c => c.IncluidoNaContagem == true).ToList().Count;

            MemoryStream stream = new MemoryStream();
            iTextSharp.text.Document docPDF = new iTextSharp.text.Document(iTextSharp.text.PageSize.A4);
            iTextSharp.text.pdf.PdfWriter writer = iTextSharp.text.pdf.PdfWriter.GetInstance(docPDF, stream);
            docPDF.Open();

            StockPriceDataSource spds = new StockPriceDataSource(context);

            #region Header
            //Declaração de variáveis
            string strEmpresa = spds.GetCompanyName();


            iTextSharp.text.pdf.PdfPTable t0 = new iTextSharp.text.pdf.PdfPTable(1);
            t0.DefaultCell.VerticalAlignment = iTextSharp.text.Element.ALIGN_MIDDLE;

            //linha1
            string strTituloDocumento = "RELATÓRIO DE CONTAGEM DE USUÁRIOS";
            iTextSharp.text.Phrase texto = new iTextSharp.text.Phrase("\n\n\n" + strTituloDocumento + " \nEMPRESA - " + strEmpresa, new iTextSharp.text.Font(iTextSharp.text.Font.FontFamily.TIMES_ROMAN, 14, iTextSharp.text.Font.BOLD, BaseColor.BLACK));
            iTextSharp.text.pdf.PdfPCell c0_1 = new iTextSharp.text.pdf.PdfPCell(texto);

            c0_1.VerticalAlignment = iTextSharp.text.Element.ALIGN_MIDDLE;
            c0_1.Border = iTextSharp.text.Rectangle.NO_BORDER;
            c0_1.HorizontalAlignment = iTextSharp.text.Element.ALIGN_CENTER;

            t0.AddCell(c0_1);
            docPDF.Add(t0);

            #endregion

            iTextSharp.text.pdf.PdfPTable t2 = new iTextSharp.text.pdf.PdfPTable(1);
            t2.DefaultCell.VerticalAlignment = iTextSharp.text.Element.ALIGN_MIDDLE;

            iTextSharp.text.Phrase texto1 = new iTextSharp.text.Phrase("\n\nÉ considerado beneficiário ativo no sistema aquele que cumpre uma ou mais das seguintes condições: \n", new iTextSharp.text.Font(iTextSharp.text.Font.FontFamily.TIMES_ROMAN, 12, iTextSharp.text.Font.NORMAL, BaseColor.BLACK));
            iTextSharp.text.pdf.PdfPCell c10 = new iTextSharp.text.pdf.PdfPCell(texto1);

            c10.VerticalAlignment = iTextSharp.text.Element.ALIGN_MIDDLE;
            c10.Border = iTextSharp.text.Rectangle.NO_BORDER;
            c10.HorizontalAlignment = iTextSharp.text.Element.ALIGN_LEFT;

            t2.AddCell(c10);

            iTextSharp.text.Phrase texto2 = new iTextSharp.text.Phrase("     a) Possuir login ativo.\n     b) Possuir contratos com status \"Em Aberto\".\n" +
                                                                        "     c) possuir alguma movimentação (exercício, resgate ou cancelamento) no ano corrente.\n", new iTextSharp.text.Font(iTextSharp.text.Font.FontFamily.TIMES_ROMAN, 12, iTextSharp.text.Font.NORMAL, BaseColor.BLACK));
            iTextSharp.text.pdf.PdfPCell c11 = new iTextSharp.text.pdf.PdfPCell(texto2);

            c11.VerticalAlignment = iTextSharp.text.Element.ALIGN_MIDDLE;
            c11.Border = iTextSharp.text.Rectangle.NO_BORDER;
            c11.HorizontalAlignment = iTextSharp.text.Element.ALIGN_LEFT;

            t2.AddCell(c11);

            iTextSharp.text.pdf.PdfPCell c12 = new iTextSharp.text.pdf.PdfPCell();
            iTextSharp.text.Paragraph p10 = new iTextSharp.text.Paragraph();
            p10.Add(new iTextSharp.text.Phrase("\nPeríodo: ", new iTextSharp.text.Font(iTextSharp.text.Font.FontFamily.TIMES_ROMAN, 12, iTextSharp.text.Font.BOLD, BaseColor.BLACK)));
            p10.Add(new iTextSharp.text.Phrase(01 + "/" + dataDocumento.AddMonths(-1).Month + "/" + dataDocumento.AddMonths(-1).Year + " a " + dataDocumento.AddDays(-1).Day + "/" + dataDocumento.AddDays(-1).Month + "/" + dataDocumento.AddDays(-1).Year, new iTextSharp.text.Font(iTextSharp.text.Font.FontFamily.TIMES_ROMAN, 12, iTextSharp.text.Font.NORMAL, BaseColor.BLACK)));
            c12.AddElement(p10);
            c12.VerticalAlignment = iTextSharp.text.Element.ALIGN_MIDDLE;
            c12.Border = iTextSharp.text.Rectangle.NO_BORDER;
            c12.HorizontalAlignment = iTextSharp.text.Element.ALIGN_LEFT;

            t2.AddCell(c12);


            iTextSharp.text.pdf.PdfPCell c14 = new iTextSharp.text.pdf.PdfPCell();
            iTextSharp.text.Paragraph p20 = new iTextSharp.text.Paragraph();
            p20.Add(new iTextSharp.text.Phrase("Número de usuários no período: ", new iTextSharp.text.Font(iTextSharp.text.Font.FontFamily.TIMES_ROMAN, 12, iTextSharp.text.Font.BOLD, BaseColor.BLACK)));
            p20.Add(new iTextSharp.text.Phrase(countColaboradoresIncluidos + "\n\n", new iTextSharp.text.Font(iTextSharp.text.Font.FontFamily.TIMES_ROMAN, 12, iTextSharp.text.Font.NORMAL, BaseColor.BLACK)));
            c14.AddElement(p20);
            c14.VerticalAlignment = iTextSharp.text.Element.ALIGN_MIDDLE;
            c14.Border = iTextSharp.text.Rectangle.NO_BORDER;
            c14.HorizontalAlignment = iTextSharp.text.Element.ALIGN_LEFT;

            t2.AddCell(c14);




            docPDF.Add(t2);

            #region Tabela Informações

            iTextSharp.text.pdf.PdfPTable t1 = new iTextSharp.text.pdf.PdfPTable(6);

            t1.HorizontalAlignment = iTextSharp.text.Element.ALIGN_CENTER;

            iTextSharp.text.pdf.PdfPCell c1_1 = new iTextSharp.text.pdf.PdfPCell();
            iTextSharp.text.Paragraph p11 = new iTextSharp.text.Paragraph(15, "Nome", new iTextSharp.text.Font(iTextSharp.text.Font.FontFamily.TIMES_ROMAN, 12, iTextSharp.text.Font.BOLD, BaseColor.BLACK));
            c1_1.AddElement(p11);
            c1_1.HorizontalAlignment = iTextSharp.text.Element.ALIGN_LEFT;
            c1_1.PaddingBottom = 5;

            iTextSharp.text.pdf.PdfPCell c1_2 = new iTextSharp.text.pdf.PdfPCell();
            iTextSharp.text.Paragraph p12 = new iTextSharp.text.Paragraph(15, "Empresa", new iTextSharp.text.Font(iTextSharp.text.Font.FontFamily.TIMES_ROMAN, 12, iTextSharp.text.Font.BOLD, BaseColor.BLACK));
            c1_2.AddElement(p12);
            c1_2.HorizontalAlignment = iTextSharp.text.Element.ALIGN_LEFT;
            c1_2.PaddingBottom = 5;


            iTextSharp.text.pdf.PdfPCell c2_1 = new iTextSharp.text.pdf.PdfPCell();
            iTextSharp.text.Paragraph p21 = new iTextSharp.text.Paragraph(15, "Status no sistema", new iTextSharp.text.Font(iTextSharp.text.Font.FontFamily.TIMES_ROMAN, 12, iTextSharp.text.Font.BOLD, BaseColor.BLACK));
            c2_1.AddElement(p21);
            c2_1.HorizontalAlignment = iTextSharp.text.Element.ALIGN_LEFT;
            c2_1.VerticalAlignment = iTextSharp.text.Element.ALIGN_TOP;
            c2_1.PaddingBottom = 5;

            iTextSharp.text.pdf.PdfPCell c2_2 = new iTextSharp.text.pdf.PdfPCell();
            iTextSharp.text.Paragraph p22 = new iTextSharp.text.Paragraph(15, "Possui lotes em aberto?", new iTextSharp.text.Font(iTextSharp.text.Font.FontFamily.TIMES_ROMAN, 12, iTextSharp.text.Font.BOLD, BaseColor.BLACK));
            c2_2.AddElement(p22);
            c2_2.HorizontalAlignment = iTextSharp.text.Element.ALIGN_LEFT;
            c2_2.PaddingBottom = 5;


            iTextSharp.text.pdf.PdfPCell c3_1 = new iTextSharp.text.pdf.PdfPCell();
            iTextSharp.text.Paragraph p31 = new iTextSharp.text.Paragraph(15, "Data da última movimentação", new iTextSharp.text.Font(iTextSharp.text.Font.FontFamily.TIMES_ROMAN, 12, iTextSharp.text.Font.BOLD, BaseColor.BLACK));
            c3_1.AddElement(p31);
            c3_1.HorizontalAlignment = iTextSharp.text.Element.ALIGN_LEFT;
            c3_1.PaddingBottom = 5;

            iTextSharp.text.pdf.PdfPCell c3_2 = new iTextSharp.text.pdf.PdfPCell();
            iTextSharp.text.Paragraph p32 = new iTextSharp.text.Paragraph(15, "Incluído na contagem?", new iTextSharp.text.Font(iTextSharp.text.Font.FontFamily.TIMES_ROMAN, 12, iTextSharp.text.Font.BOLD, BaseColor.BLACK));
            c3_2.AddElement(p32);
            c3_2.HorizontalAlignment = iTextSharp.text.Element.ALIGN_LEFT;
            c3_2.PaddingBottom = 5;


            t1.AddCell(c1_1);
            t1.AddCell(c1_2);
            t1.AddCell(c2_1);
            t1.AddCell(c2_2);
            t1.AddCell(c3_1);
            t1.AddCell(c3_2);


            foreach (ColaboradorSituacao colaborador in lstColaboradores)
            {
                string incluidoContagem = colaborador.IncluidoNaContagem ? "Sim" : "Não";
                string loteAberto = colaborador.LoteEmAberto ? "Sim" : "Não";
                string dataMovimentacao = "-";
                if (colaborador.UltimaMovimentacao != null)
                {
                    dataMovimentacao = colaborador.UltimaMovimentacao.Value.ToShortDateString();
                }

                iTextSharp.text.pdf.PdfPCell c4_1 = new iTextSharp.text.pdf.PdfPCell();
                iTextSharp.text.Paragraph p41 = new iTextSharp.text.Paragraph(15, colaborador.EmployeeName, new iTextSharp.text.Font(iTextSharp.text.Font.FontFamily.TIMES_ROMAN, 12, iTextSharp.text.Font.NORMAL, BaseColor.BLACK));
                c4_1.AddElement(p41);
                c4_1.HorizontalAlignment = iTextSharp.text.Element.ALIGN_LEFT;
                c4_1.PaddingBottom = 5;

                t1.AddCell(c4_1);

                iTextSharp.text.pdf.PdfPCell c4_2 = new iTextSharp.text.pdf.PdfPCell();
                iTextSharp.text.Paragraph p42 = new iTextSharp.text.Paragraph(15, colaborador.Empresa, new iTextSharp.text.Font(iTextSharp.text.Font.FontFamily.TIMES_ROMAN, 12, iTextSharp.text.Font.NORMAL, BaseColor.BLACK));
                c4_2.AddElement(p42);
                c4_2.HorizontalAlignment = iTextSharp.text.Element.ALIGN_LEFT;
                c4_2.PaddingBottom = 5;

                t1.AddCell(c4_2);

                iTextSharp.text.pdf.PdfPCell c4_3 = new iTextSharp.text.pdf.PdfPCell();
                iTextSharp.text.Paragraph p43 = new iTextSharp.text.Paragraph(15, colaborador.StatusSistema, new iTextSharp.text.Font(iTextSharp.text.Font.FontFamily.TIMES_ROMAN, 12, iTextSharp.text.Font.NORMAL, BaseColor.BLACK));
                c4_3.AddElement(p43);
                c4_3.HorizontalAlignment = iTextSharp.text.Element.ALIGN_LEFT;
                c4_3.PaddingBottom = 5;

                t1.AddCell(c4_3);

                iTextSharp.text.pdf.PdfPCell c4_4 = new iTextSharp.text.pdf.PdfPCell();
                iTextSharp.text.Paragraph p44 = new iTextSharp.text.Paragraph(15, loteAberto, new iTextSharp.text.Font(iTextSharp.text.Font.FontFamily.TIMES_ROMAN, 12, iTextSharp.text.Font.NORMAL, BaseColor.BLACK));
                c4_4.AddElement(p44);
                c4_4.HorizontalAlignment = iTextSharp.text.Element.ALIGN_LEFT;
                c4_4.PaddingBottom = 5;

                t1.AddCell(c4_4);

                iTextSharp.text.pdf.PdfPCell c4_5 = new iTextSharp.text.pdf.PdfPCell();
                iTextSharp.text.Paragraph p45 = new iTextSharp.text.Paragraph(15, dataMovimentacao, new iTextSharp.text.Font(iTextSharp.text.Font.FontFamily.TIMES_ROMAN, 12, iTextSharp.text.Font.NORMAL, BaseColor.BLACK));
                c4_5.AddElement(p45);
                c4_5.HorizontalAlignment = iTextSharp.text.Element.ALIGN_LEFT;
                c4_5.PaddingBottom = 5;

                t1.AddCell(c4_5);

                iTextSharp.text.pdf.PdfPCell c4_6 = new iTextSharp.text.pdf.PdfPCell();
                iTextSharp.text.Paragraph p46 = new iTextSharp.text.Paragraph(15, incluidoContagem, new iTextSharp.text.Font(iTextSharp.text.Font.FontFamily.TIMES_ROMAN, 12, iTextSharp.text.Font.NORMAL, BaseColor.BLACK));
                c4_6.AddElement(p46);
                c4_6.HorizontalAlignment = iTextSharp.text.Element.ALIGN_LEFT;
                c4_6.PaddingBottom = 5;

                t1.AddCell(c4_6);

            }

            t1.HorizontalAlignment = iTextSharp.text.Element.ALIGN_CENTER;
            float[] totalWidths0 = new float[] { 180, 110, 50, 50, 80, 60 };
            t1.SetTotalWidth(totalWidths0);
            t1.LockedWidth = true;
            docPDF.Add(t1);
            #endregion

            writer.CloseStream = false;
            docPDF.Close();

            return stream;
        }
        #endregion
        #region Thread para futuros testes
        public void Thread()
        {
            //Thread newThread = new Thread(() => { listaColaboradores = ConstruirListaColaboradores(DateTime.Now, listaColaboradores, connStr); });
            //newThread.Start();
        }
        #endregion
        #endregion
    }
}