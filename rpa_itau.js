const puppeteer = require("puppeteer");
const XLSX = require("xlsx");
const path = require("path");

/**
 * Inicializa o navegador automatizado
 */
async function iniciarBrowser() {
  return puppeteer.launch({
    headless: false,
    defaultViewport: null
  });
}

/**
 * Acessa a página do ativo e aguarda o carregamento dos dados
 */
async function carregarPagina(page, url) {
  await page.goto(url, { waitUntil: "networkidle2" });

  // Aguarda até que os dados financeiros estejam visíveis
  await page.waitForFunction(() =>
    document.body.innerText.includes("Cotação")
  );
}

/**
 * Extrai indicadores financeiros da página
 */
async function extrairIndicadores(page) {
  return page.evaluate(() => {
    const linhas = document.body.innerText.split("\n");

    const buscarValor = (indicador) => {
      const index = linhas.findIndex(l =>
        l.toLowerCase().includes(indicador.toLowerCase())
      );
      return index !== -1 ? linhas[index + 1] : "N/A";
    };

    return {
      ativo: "ITUB4",
      cotacao: buscarValor("Cotação"),
      dividendYield: buscarValor("Dividend Yield"),
      pl: buscarValor("P/L"),
      pvp: buscarValor("P/VP")
    };
  });
}

/**
 * Gera um arquivo Excel com os dados coletados
 */
function gerarExcel(dados) {
  const workbook = XLSX.utils.book_new();
  const sheet = XLSX.utils.json_to_sheet([dados]);

  XLSX.utils.book_append_sheet(workbook, sheet, "ITAU");

  const arquivo = path.join(process.cwd(), "relatorio_itau.xlsx");
  XLSX.writeFile(workbook, arquivo);

  return arquivo;
}

/**
 * Fluxo principal do RPA
 */
async function executar() {
  console.log("Iniciando RPA Itaú...");

  const browser = await iniciarBrowser();
  const page = await browser.newPage();

  await carregarPagina(page, "https://investidor10.com.br/acoes/itub4/");

  const dados = await extrairIndicadores(page);
  console.log("Dados coletados:", dados);

  const caminhoArquivo = gerarExcel(dados);
  console.log("Relatório salvo em:", caminhoArquivo);

  await browser.close();
  console.log("Processo finalizado.");
}

// Ponto de entrada da aplicação
executar();
