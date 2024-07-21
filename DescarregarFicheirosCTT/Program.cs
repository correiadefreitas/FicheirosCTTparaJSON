// See https://aka.ms/new-console-template for more information
using INIParser;
using OpenQA.Selenium.Chrome;
using OpenQA.Selenium;
using System.IO.Compression;
using System.Text.Json;
using System.Collections.ObjectModel;
using System.Text.Encodings.Web;
using System.Text.Unicode;
using System.Text;

if (args.Count() == 0 ||
    !string.IsNullOrEmpty(args.FirstOrDefault(x => x.Equals("-ajuda", StringComparison.CurrentCultureIgnoreCase) 
    || x.Equals("/h", StringComparison.CurrentCultureIgnoreCase) 
    || x.Equals("/help", StringComparison.CurrentCultureIgnoreCase) 
    || x.Equals("-a", StringComparison.CurrentCultureIgnoreCase))))
{
    Console.WriteLine("Programa que descarrega, descompacta e converte para JSON os ficheiros dos CTT relativos aos códigos postais.\r\n\tParametros:" +
        "\r\n\t\t-obter\tUtiliza as seguintes entradas do 'FicheirosCTTparaJSON.ini'" +
        "\r\n\t\t\t[obter] nome_utilizador Nome de utilizador utilizado no site dos CTT (Ex.: o_seu_email@empresa.pt)" +
        "\r\n\t\t\t[obter] senha Senha utilizada no site dos CTT" +
        "\r\n\t\t\t[obter] morada_autenticacao URL de autenticação do site dos CTT (Ex.: https://www.ctt.pt/fecas/login)" +
        "\r\n\t\t\t[obter] morada_zip URL do ZIP com os ficheiros dos códigos postais (Ex.: https://appserver2.ctt.pt/feapl_2/app/restricted/postalCodeSearch/postalCodeDownloadFiles!downloadPostalCodeFile.jspx)" +
        "\r\n\t\t\t[obter] nome_ficheiro_ctt Nome completo do ZIP com os ficheiros dos códigos postais (todos_cp.zip)" +
        "\r\n\t\t-extrair\tUtiliza as seguintes entradas do 'FicheirosCTTparaJSON.ini'" +
        "\r\n\t\t\t[extrair] nome_utilizador Nome de utilizador utilizado no site dos CTT" +
        "\r\n\t\t-json\t");
    return;
}

// Lê as credenciais e outras informações do ficheiro 'FicheirosCTTparaJSON.ini'
var data = new IniFile("FicheirosCTTparaJSON.ini");

string username = data["obter", "nome_utilizador"] ?? throw new Exception("Entrada ausente no ficheiro de configuração 'FicheirosCTTparaJSON.ini': [obter] -> nome_utilizador");
string password = data["obter", "senha"] ?? throw new Exception("Entrada ausente no ficheiro de configuração 'FicheirosCTTparaJSON.ini': [obter] -> senha");
string morada_autenticacao = data["obter", "morada_autenticacao"] ?? throw new Exception("Entrada ausente no ficheiro de configuração 'FicheirosCTTparaJSON.ini': [obter] -> morada_autenticacao");
string morada_zip = data["obter", "morada_zip"] ?? throw new Exception("Entrada ausente no ficheiro de configuração 'FicheirosCTTparaJSON.ini': [obter] -> morada_zip");
string nome_ficheiro_ctt = data["obter", "nome_ficheiro_ctt"] ?? throw new Exception("Entrada ausente no ficheiro de configuração 'FicheirosCTTparaJSON.ini': [obter] -> nome_ficheiro_ctt");

string nome_pasta_ctt = data["extrair", "nome_pasta_ctt"] ?? throw new Exception("Entrada ausente no ficheiro de configuração 'FicheirosCTTparaJSON.ini': [extrair] -> nome_pasta_ctt");

string nome_ficheiro_distritos = data["json", "nome_ficheiro_distritos"] ?? throw new Exception("Entrada ausente no ficheiro de configuração 'FicheirosCTTparaJSON.ini': [json] -> nome_ficheiro_distritos");
string nome_ficheiro_concelhos = data["json", "nome_ficheiro_concelhos"] ?? throw new Exception("Entrada ausente no ficheiro de configuração 'FicheirosCTTparaJSON.ini': [json] -> nome_ficheiro_concelhos");
string nome_ficheiro_todos_cp = data["json", "nome_ficheiro_todos_cp"] ?? throw new Exception("Entrada ausente no ficheiro de configuração 'FicheirosCTTparaJSON.ini': [json] -> nome_ficheiro_todos_cp");
string nome_pasta_json = data["json", "nome_pasta_json"] ?? throw new Exception("Entrada ausente no ficheiro de configuração 'FicheirosCTTparaJSON.ini': [json] -> nome_pasta_json");


if (!string.IsNullOrEmpty(args.FirstOrDefault(x => x.Equals("-obter",StringComparison.CurrentCultureIgnoreCase))) && !File.Exists(nome_ficheiro_ctt)) 
{

    var thisOptions = new ChromeOptions();
    thisOptions.AddArgument("--headless");

    thisOptions.AddArgument("--user-agent=Mozilla/5.0 (Windows NT 6.1; WOW64; Trident/7.0; rv:11.0) like Gecko");
    thisOptions.AddArgument("--log-level=3"); // Desativa o log do ChromeDriver

    // Configurações de download
    string downloadDirectory = Directory.GetCurrentDirectory();
    thisOptions.AddUserProfilePreference("download.default_directory", downloadDirectory);

    // Inicializa o browser
    var browser = new ChromeDriver(thisOptions);

    // Fazer o login
    browser.Navigate().GoToUrl(morada_autenticacao);
    Thread.Sleep(1000);

    // Modal de Aceitar cookies
    browser.FindElement(By.Id("onetrust-accept-btn-handler")).Click();
    Thread.Sleep(1000);

    // Login
    browser.FindElement(By.Id("username")).SendKeys(username);
    Thread.Sleep(1000);
    browser.FindElement(By.Id("password")).SendKeys(password);
    Thread.Sleep(1000);
    browser.FindElement(By.Name("btnSubmit")).Click();
    Thread.Sleep(1000);

    // Ir ao URL do ficheiro directo
    browser.Navigate().GoToUrl(morada_zip);

    // Espera pelo download do ficheiro
    int timer = 0;
    while (!File.Exists(nome_ficheiro_ctt) && timer < 30)
    {
        Thread.Sleep(1000);
        timer += 1;
    }

    // Fecha o browser
    browser.Quit();
}

if (!string.IsNullOrEmpty(args.FirstOrDefault(x => x.Equals("-extrair", StringComparison.CurrentCultureIgnoreCase))) && File.Exists(nome_ficheiro_ctt))
{
    try
    {
        // Descompacta 
        Directory.CreateDirectory(nome_pasta_ctt);
        ZipFile.ExtractToDirectory(nome_ficheiro_ctt, nome_pasta_ctt, true);
    }
    catch (Exception)
    {
        Console.WriteLine($"Erro a descompactar {nome_ficheiro_ctt} para a pasta {nome_ficheiro_ctt}");
        return;
    }

}

if (!string.IsNullOrEmpty(args.FirstOrDefault(x => x.Equals("-json", StringComparison.CurrentCultureIgnoreCase))))
{
    IList<String> cabecalhoDistritos = new ReadOnlyCollection<string>([
        "DD",
        "DESIG"
     ]);
    IList<String> cabecalhoConcelhos = new ReadOnlyCollection<string>([
        "DD",
        "CC",
        "DESIG"
    ]);
    IList<String> cabecalhoTodosCP = new ReadOnlyCollection<string>([
        "DD",
        "CC",
        "LLLL",
        "LOCALIDADE",
        "ART_COD",
        "ART_TIPO",
        "PRI_PREP",
        "ART_TITULO",
        "SEG_PREP",
        "ART_DESIG",
        "ART_LOCAL",
        "TROÇO",
        "PORTA",
        "CLIENTE",
        "CP4",
        "CP3",
        "CPALF"
    ]);

    Directory.CreateDirectory(nome_pasta_json);

    ConverterTxtEmJson(nome_pasta_ctt, nome_ficheiro_distritos, nome_pasta_json, cabecalhoDistritos);
    ConverterTxtEmJson(nome_pasta_ctt, nome_ficheiro_concelhos, nome_pasta_json, cabecalhoConcelhos);
    ConverterTxtEmJson(nome_pasta_ctt, nome_ficheiro_todos_cp, nome_pasta_json, cabecalhoTodosCP);
}


static void ConverterTxtEmJson(string nome_pasta_ctt, string nome_ficheiro_txt, string nome_pasta_json, IList<string> cabecalho)
{
    string[] csvLines;
    try
    {
        // Ler Dados Distritos
        csvLines = File.ReadAllLines("./" + nome_pasta_ctt + "/" + nome_ficheiro_txt, System.Text.Encoding.Latin1);
    }
    catch (Exception ex)
    {
        Console.WriteLine($"Erro a ler {nome_ficheiro_txt} da pasta {nome_pasta_ctt}:\r\n{ex}");
        return;
    }

    string jsonString;
    try
    {
        // Processar CVS Distritos
        var jsonData = csvLines.Select(line =>
        {
            var values = line.Split(';');
            var obj = new Dictionary<string, string>();
            for (int i = 0; i < cabecalho.Count && i < values.Length; i++)
            {
                obj[cabecalho[i]] = Encoding.UTF8.GetString(Encoding.Default.GetBytes(values[i]));
            }
            return obj;
        }).ToList();

        // Serializar Distritos para JSON
        var options = new JsonSerializerOptions { WriteIndented = true, Encoder = JavaScriptEncoder.Create(UnicodeRanges.BasicLatin, UnicodeRanges.Latin1Supplement) };
        jsonString = JsonSerializer.Serialize(jsonData, options);
    }
    catch (Exception ex)
    {
        Console.WriteLine($"Erro a serializar {nome_ficheiro_txt} da pasta {nome_pasta_ctt}:\r\n{ex}");
        return;
    }

    try
    {
        // Gravar JSON Distritos
        File.WriteAllText("./" + nome_pasta_json + "/" + nome_ficheiro_txt.Replace(".txt", ".json"), jsonString);
    }
    catch (Exception ex)
    {
        Console.WriteLine($"Erro a descompactar {nome_ficheiro_txt} para a pasta {nome_pasta_json}:\r\n{ex}");
        return;
    }
}