const puppeteer = require("puppeteer");
const ExcelJS = require('exceljs');

console.log("aplicação rodando");
const today = new Date();

async function getDolarToday() {

    const navegador = await puppeteer.launch({ headless: false });
    const paginaWeb = await navegador.newPage();

    const moedaBase = 'real';
    const moedaConversao = 'dolar';

    const url = `https://www.google.com/search?q=${moedaConversao}+para+${moedaBase}`

    await paginaWeb.goto(url);

    const resultadoConversao = await paginaWeb.$eval('span.DFlfde.SwHCTb', (element) => {
        return element.textContent;
    });

    await navegador.close();

    return resultadoConversao;
}

async function criarPlanilha() {
    const nomeArquivo = 'resultado_conversao.xlsx';
    const planilha = new ExcelJS.Workbook();

    try {
        await planilha.xlsx.readFile(nomeArquivo);
    } catch (error) {
        console.log('Arquivo não encontrado. Criando uma nova planilha.');
        const novaPlanilha = new ExcelJS.Workbook();
        const novaWorksheet = novaPlanilha.addWorksheet('Sheet 1');
        await novaPlanilha.xlsx.writeFile(nomeArquivo);
        console.log('Nova planilha criada com sucesso.');
        return;
    }

    const worksheet = planilha.getWorksheet('Sheet 1');

    let ultimaLinha = 1;

    worksheet.getColumn('B').eachCell({ includeEmpty: false }, (cell, rowNumber) => {
        ultimaLinha = rowNumber;
    });

    const novaLinha = ultimaLinha + 1;

    const resultadoConversao = await getDolarToday();

    if (novaLinha === 1) {
        worksheet.getCell(`A1`).value = today;
        worksheet.getCell(`B1`).value = resultadoConversao;

    } else {
        worksheet.getCell(`A${novaLinha}`).value = today;
        worksheet.getCell(`B${novaLinha}`).value = resultadoConversao;
    }

    await planilha.xlsx.writeFile(nomeArquivo);
    console.log('Planilha atualizada!');

};

criarPlanilha();