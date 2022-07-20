using iTextSharp.text.pdf;
using iTextSharp.text.pdf.parser;
using Microsoft.Extensions.Configuration;
using MySql.Data.MySqlClient;
using System.Text;
using System.IO.Compression;

internal class Class
{
    static void Main(string[] args)
    {
        var builder = new ConfigurationBuilder()
               .AddJsonFile($"appsettings.json", true, true);

        var config = builder.Build();
        //Caminho do Log
        var log = config["SME_Log_Certificados"];

        //Criando Log
        StringBuilder sb = new StringBuilder();
        StringBuilder sbErro = new StringBuilder();
        sb.AppendLine("Iniciando o processo");

        //Passando o diretório via Variável de ambiente
        var arquivos = config["SME_Caminho_Arquivos"];

        //Diretório armazenamento e leitura de arquivos


        string[] files = Directory.GetFiles(arquivos); //paginar essa consulta e pegar de 5mil em 5mil arquivos


        //Diretório onde os arquivos ".msg" se encontram
	//Console.WriteLine("Acessando o diretório ...");

        sb.AppendLine("Acessou o diretório");
        var connetionString = config["SME_Certificados_Cs"];

        using MySqlConnection cn = new(connetionString);
        cn.Open();

        foreach (var file in files)
        {
            DateTime modifyTime = File.GetLastWriteTime(file);

            string dataModificacao = modifyTime.Year.ToString();

            try
            {
                if (dataModificacao == "2022" || dataModificacao == "2021")
                {
                    using var msg = new MsgReader.Outlook.Storage.Message(file);
                    Encoding.RegisterProvider(System.Text.CodePagesEncodingProvider.Instance);
                    //Faço uma verificação nos arquivos, se não seguir as regras de nomenclatura, já passo direto.                   
                    //Exemplo em que não é encontrado o RF: Curso realizado na SME -_C12017430718_I190679
                    //Exemplo em que possuem números com pontuação: Curso realizado na SME -014.261.29_C1HOM21182_I324370


                    //Campos que não acessam os Anexos

                    var assuntoEmail = msg.Subject;
                    var htmlBody = msg.BodyHtml;
                    var nomeArquivo = msg.GetAttachmentNames();

                    string rfUsuario = nomeArquivo.Substring(0, 7);
                    string arquivoNome = assuntoEmail.Substring(assuntoEmail.LastIndexOf('-') + 1);
                    string firstCharacter = arquivoNome.Substring(0, 2);

                    //Número de Homologação
                    int hFrom = arquivoNome.IndexOf("_") + "_".Length;
                    int hTo = arquivoNome.LastIndexOf("_");

                    string numHomologacaoCurso = arquivoNome.Substring(hFrom, hTo - hFrom);

                    string firstCharacterNumHom = numHomologacaoCurso.Substring(0, 1);

                    if (firstCharacterNumHom == "_")
                    {
                        numHomologacaoCurso = numHomologacaoCurso.Replace(firstCharacterNumHom, "");
                    }

                    sb.AppendLine("Verificou se existe caso para o Rf: " + rfUsuario);

                    //Preciso verificar se os casos já existem na tabela

                    //lendo os sAnexos
                    foreach (MsgReader.Outlook.Storage.Attachment itmAttachment in msg.Attachments)
                    {

                        //Convertendo para base 64
                        var oData = itmAttachment.Data;
                        var arquivoString = Convert.ToBase64String(oData);

                        using var compressedStream = new MemoryStream();
                        using var compressor = new GZipStream(compressedStream, CompressionLevel.SmallestSize, leaveOpen: true);
                        compressor.Write(oData, 0, oData.Length);
                        compressor.Close();
                        var compressedArray = compressedStream.ToArray();
                        var anexo = Convert.ToBase64String(compressedArray);

                        sb.AppendLine("Iniciou a verificação do Rf: " + rfUsuario);
                        //Se conter pontuação nem efetuo as validações ou se iniciar com _
                        if (firstCharacter != " _" || firstCharacter != "_")
                        {
                            if (!arquivoNome.Contains('.'))
                            {
                                //Iniciando a leitura dos arquivos
                                using PdfReader reader = new(Convert.FromBase64String(arquivoString));

                                LerPdf(reader, out var nomeCurso, out var dataConclusaoCurso);


                                //Só vai inserir caso não exista na tabela
                                if (VerificarSeJaExiste(numHomologacaoCurso, rfUsuario, cn) == 0)
                                {
                                    //insert - debito tecnico salvar arquivos em uma pasta e colocar o path no banco
                                    InsertPdfDatabase(rfUsuario, numHomologacaoCurso, cn, anexo, nomeCurso, dataConclusaoCurso, Guid.NewGuid());

                                    //Console.WriteLine("Informação inserida com sucesso para o RF: " + rfUsuario);
                                    sb.AppendLine("Inseriu no banco o Rf: " + rfUsuario);
                                }
                            }
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                sbErro.AppendLine("Erro no arquivo: " + file + " Descrição:" + ex.Message);
                File.AppendAllText(log + "log_ERRO" + DateTime.UtcNow.ToString("ddMMyyyyHHmmss") + ".txt", sbErro.ToString());
            }

        }

        cn.Close();

        Console.WriteLine("Fim do processo!!");

        sb.AppendLine("Processo executado com sucesso");

        File.AppendAllText(log + "log_" + DateTime.UtcNow.ToString("ddMMyyyyHHmmss") + ".txt", sb.ToString());
        File.AppendAllText(log + "log_ERRO" + DateTime.UtcNow.ToString("ddMMyyyyHHmmss") + ".txt", sbErro.ToString());
    }

    private static void LerPdf(PdfReader reader, out string nomeCurso, out string dataConclusaoCurso)
    {
        nomeCurso = string.Empty;
        dataConclusaoCurso = string.Empty;
        for (int pageNo = 1; pageNo <= 1; pageNo++)
        {
            ITextExtractionStrategy strategy = new SimpleTextExtractionStrategy();
            string textPdf = PdfTextExtractor.GetTextFromPage(reader, pageNo, strategy);
            textPdf = Encoding.UTF8.GetString(ASCIIEncoding.Convert(Encoding.Default, Encoding.UTF8, Encoding.Default.GetBytes(textPdf)));

            //Se conter PARTICIPOU tratarei nesse if
            if (textPdf.Contains("participou") && textPdf.Contains("promovido"))
            {
                int pFrom = textPdf.IndexOf("participou") + "participou".Length;

                int pTo = textPdf.LastIndexOf("promovido");

                //Validação para casos encontrados em teste
                nomeCurso = textPdf[pFrom..pTo];
            }

            //Localizando a data de conclusão
            int dcFrom = textPdf.IndexOf("São Paulo, ") + "São Paulo, ".Length;
            int dcTo = textPdf.LastIndexOf(".\nPMSP");
            if (dcTo == -1)
            {
                dcTo = textPdf.LastIndexOf(".\nC E R T I F I C A D O");
                if (dcTo == -1)
                {
                    dcTo = textPdf.LastIndexOf(".\n                  Atestamos");
                }
            }

            dataConclusaoCurso = RetornaMes(textPdf[dcFrom..dcTo].Replace(" de ", "/").Replace(" ", ""));
        }
    }

    private static int VerificarSeJaExiste(string numHomologacaoCurso, string numeroRf, MySqlConnection cn)
    {
        //Preciso verificar se os casos já existem na tabela
        var query = "SELECT COUNT(*) FROM tb_arquivo_certificado WHERE num_homolog_curso = @num_homolog_curso and rf = @numeroRf";
        MySqlCommand cmd = new(query, cn);
        cmd.Parameters.AddWithValue("@num_homolog_curso", numHomologacaoCurso);
        cmd.Parameters.AddWithValue("@numeroRf", numeroRf);
        return Convert.ToInt32(cmd.ExecuteScalar());
    }

    private static void InsertPdfDatabase(string rfUsuario, string numHomologacaoCurso, MySqlConnection cn, string anexo, string nomeCurso, string dataConclusaoCurso, Guid id)
    {
        string queryInsert = "INSERT INTO tb_arquivo_certificado(id,rf,num_homolog_curso,nome_curso,arquivo,dt_conclusao,dt_execucao) VALUES(@id,@rf,@num_homolog_curso,@nome_curso,@arquivo,@dt_conclusao,@dt_execucao)";
        using (MySqlCommand cmdInsert = new MySqlCommand(queryInsert, cn))
        {
            cmdInsert.Parameters.Add("@id", MySqlDbType.VarChar).Value = id;
            cmdInsert.Parameters.Add("@rf", MySqlDbType.VarChar).Value = rfUsuario;
            cmdInsert.Parameters.Add("@num_homolog_curso", MySqlDbType.VarChar).Value = numHomologacaoCurso;
            cmdInsert.Parameters.Add("@nome_curso", MySqlDbType.VarChar).Value = nomeCurso;
            cmdInsert.Parameters.Add("@arquivo", MySqlDbType.MediumText).Value = anexo;
            cmdInsert.Parameters.Add("@dt_conclusao", MySqlDbType.VarChar).Value = dataConclusaoCurso;
            cmdInsert.Parameters.AddWithValue("@dt_execucao", DateTime.Now);
            cmdInsert.ExecuteNonQuery();
        }
    }

    public static string RetornaMes(string data)
    {
        //Quebrando as datas
        string dataTratada = "";
        string mesAux = "";

        //Dia
        string dia = data.Split('/')[0];

        //Mês
        int mesFrom = data.IndexOf("/") + "/".Length;
        int mesTo = data.LastIndexOf("/");
        string mes = data.Substring(mesFrom, mesTo - mesFrom);

        //Ano
        string ano = data.Split('/')[2];

        switch (mes)
        {
            case "janeiro":
                mesAux = "01";
                break;
            case "fevereiro":
                mesAux = "02";
                break;
            case "março":
                mesAux = "03";
                break;
            case "abril":
                mesAux = "04";
                break;
            case "maio":
                mesAux = "05";
                break;
            case "junho":
                mesAux = "06";
                break;
            case "julho":
                mesAux = "07";
                break;
            case "agosto":
                mesAux = "08";
                break;
            case "setembro":
                mesAux = "09";
                break;
            case "outubro":
                mesAux = "10";
                break;
            case "novembro":
                mesAux = "11";
                break;
            case "dezembro":
                mesAux = "12";
                break;
            default:
                Console.Write("Mês inválido....\n");
                break;
        }

        dataTratada = dia + "/" + mesAux + "/" + ano;

        return dataTratada;
    }
}
