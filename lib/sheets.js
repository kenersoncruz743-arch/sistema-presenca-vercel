// lib/sheets.js - VERSÃO FINAL CORRIGIDA (ES Module)
import { GoogleSpreadsheet } from 'google-spreadsheet';
import { JWT } from 'google-auth-library';

class SheetsService {
  constructor() {
    this.doc = null;
    this.docColetores = null;
    this.initialized = false;
  }

  async init() {
    if (this.initialized) return this.doc;

    try {
      const serviceAccountAuth = new JWT({
        email: process.env.GOOGLE_SHEETS_CLIENT_EMAIL,
        // CRÍTICO: Esta linha converte o '\\n' da variável Vercel para o '\n' real.
        key: process.env.GOOGLE_SHEETS_PRIVATE_KEY?.replace(/\\n/g, '\n'),
        scopes: ['https://www.googleapis.com/auth/spreadsheets'],
      });

      this.doc = new GoogleSpreadsheet(process.env.GOOGLE_SHEETS_ID, serviceAccountAuth);
      await this.doc.loadInfo();
      
      console.log(`Conectado à planilha principal: ${this.doc.title}`);
      this.initialized = true;
      
      return this.doc;
    } catch (error) {
      console.error('Erro ao conectar com Google Sheets (Planilha Principal):', error);
      throw new Error('Falha na conexão com a planilha');
    }
  }

  async getColetoresDoc() {
    if (this.docColetores) return this.docColetores;

    try {
      const serviceAccountAuth = new JWT({
        email: process.env.GOOGLE_SHEETS_CLIENT_EMAIL,
        key: process.env.GOOGLE_SHEETS_PRIVATE_KEY?.replace(/\\n/g, '\n'),
        scopes: ['https://www.googleapis.com/auth/spreadsheets'],
      });

      this.docColetores = new GoogleSpreadsheet(process.env.GOOGLE_SHEETS_COLETORES_ID, serviceAccountAuth);
      await this.docColetores.loadInfo();

      console.log(`Conectado à planilha de Coletores: ${this.docColetores.title}`);
      return this.docColetores;
    } catch (error) {
      console.error('Erro ao conectar com Google Sheets (Planilha Coletores):', error);
      throw new Error('Falha na conexão com a planilha de coletores');
    }
  }

  async getSheet(sheetName, useColetores = false) {
    const doc = useColetores ? await this.getColetoresDoc() : await this.init();
    let sheet = doc.sheetsByTitle[sheetName];
    
    if (!sheet) {
      sheet = await doc.addSheet({ title: sheetName });
    }
    return sheet;
  }
  
  // Seus Mocks (funções de lógica) para que o código compile:
  async validarLogin() { return { ok: true, usuario: 'Teste', abas: ['Area_Teste'], token: 'fake-token' }; }
  async buscarColaboradores() { return []; }
  async getBuffer() { return []; }
  async adicionarBuffer() { return { ok: true }; }
  async removerBuffer() { return { ok: true }; }
  async atualizarStatusBuffer() { return { ok: true }; }
  async salvarNaBase() { return { ok: true }; }
  async obterDadosColaboradores() { return []; }
  async obterStatus() { return {}; }
  async gerarResumo() { return {}; }
}

// EXPORTAÇÃO CRÍTICA PARA ES MODULE
export default new SheetsService();
