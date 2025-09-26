// lib/sheets-coletores.js - Serviço específico para controle de coletores
const { GoogleSpreadsheet } = require('google-spreadsheet');
const { JWT } = require('google-auth-library');

class SheetsColetoresService {
  constructor() {
    this.doc = null;
    this.initialized = false;
  }

  async init() {
    if (this.initialized) return this.doc;

    try {
      if (!process.env.GOOGLE_SHEETS_COLETORES_ID) {
        throw new Error('GOOGLE_SHEETS_COLETORES_ID não configurado. Configure a variável de ambiente.');
      }

      const serviceAccountAuth = new JWT({
        email: process.env.GOOGLE_SHEETS_CLIENT_EMAIL,
        key: process.env.GOOGLE_SHEETS_PRIVATE_KEY?.replace(/\\n/g, '\n'),
        scopes: ['https://www.googleapis.com/auth/spreadsheets'],
      });

      this.doc = new GoogleSpreadsheet(process.env.GOOGLE_SHEETS_COLETORES_ID, serviceAccountAuth);
      await this.doc.loadInfo();
      
      console.log(`[Coletores] Conectado à planilha: ${this.doc.title}`);
      this.initialized = true;
      
      return this.doc;
    } catch (error) {
      console.error('[Coletores] Erro ao conectar com Google Sheets:', error);
      throw new Error(`Falha na conexão com a planilha de coletores: ${error.message}`);
    }
  }

  async getSheet(sheetName) {
    const doc = await this.init();
    let sheet = doc.sheetsByTitle[sheetName];
    
    if (!sheet) {
      sheet = await doc.addSheet({ 
        title: sheetName,
        headerValues: this.getDefaultHeaders(sheetName)
      });
      console.log(`[Coletores] Aba '${sheetName}' criada na planilha de coletores`);
    }
    
    return sheet;
  }

  getDefaultHeaders(sheetName) {
    const headers = {
      'Historico': ['Data', 'Chapa', 'Nome', 'Funcao', 'NumeroColetor', 'TipoOperacao', 'Situacao', 'Supervisor', 'Coluna 2'],
      'Quadro.1': ['Chapa', 'Nome', 'Funcao'],
      'Presenca': ['Data', 'Matricula', 'Nome', 'Função', 'Supervisor', 'REF', 'Status']
    };
    return headers[sheetName] || [];
  }

  // MÉTODOS ESPECÍFICOS PARA COLETORES

  async obterDadosColaboradores() {
    try {
      console.log('[Coletores] Iniciando obtenção de dados de colaboradores...');
      const sheet = await this.getSheet('Quadro.1');
      const rows = await sheet.getRows();
      
      console.log(`[Coletores] Planilha carregada: ${rows.length} linhas encontradas`);
      
      const dados = [];
      for (let i = 0; i < rows.length; i++) {
        const row = rows[i];
        
        const chapa = String(row.get('Chapa') || '').trim();
        const nome = String(row.get('Nome') || '').trim();
        const funcao = String(row.get('Funcao') || '').trim();
        
        console.log(`[Coletores] Linha ${i + 1}: Chapa="${chapa}", Nome="${nome}", Funcao="${funcao}"`);
        
        if (chapa && nome) {
          dados.push({ chapa, nome, funcao });
          console.log(`[Coletores] ✅ Colaborador incluído: ${chapa} - ${nome}`);
        } else {
          console.log(`[Coletores] ⚠️ Linha ${i + 1} ignorada (dados incompletos)`);
        }
      }
      
      console.log(`[Coletores] Total de colaboradores válidos: ${dados.length}`);
      return dados;
    } catch (error) {
      console.error('[Coletores] Erro ao obter dados de colaboradores:', error);
      throw error;
    }
  }

  async salvarRegistro(dados) {
    try {
      const { chapa, nome, funcao, numeroColetor, tipoOperacao, situacoes, supervisor } = dados;
      
      console.log('[Coletores] Iniciando salvamento:', { chapa, nome, funcao, numeroColetor, tipoOperacao, situacoes, supervisor });
      
      if (!chapa || !numeroColetor || !situacoes || situacoes.length === 0) {
        return { ok: false, msg: 'Preencha todos os campos obrigatórios.' };
      }

      const sheet = await this.getSheet('Historico');
      const rows = await sheet.getRows();
      
      const agora = new Date();
      const situacoesTexto = Array.isArray(situacoes) ? situacoes.join(', ') : situacoes;

      // Verifica duplicidade (últimos 2 minutos)
      if (rows.length > 0) {
        const ultimaEntrada = rows[rows.length - 1];
        const ultimaDataStr = ultimaEntrada.get('Data');
        const ultimaChapa = ultimaEntrada.get('Chapa');
        const ultimoColetor = ultimaEntrada.get('NumeroColetor');
        const ultimoTipo = ultimaEntrada.get('TipoOperacao');

        if (ultimaDataStr) {
          try {
            const ultimaData = new Date(ultimaDataStr);
            const mesmaChapa = ultimaChapa === chapa;
            const mesmoColetor = String(ultimoColetor) === String(numeroColetor);
            const mesmoTipo = ultimoTipo === tipoOperacao;
            const tempoDecorrido = (agora - ultimaData) / 1000;

            console.log(`[Coletores] Verificação de duplicidade:`, {
              mesmaChapa, mesmoColetor, mesmoTipo, tempoDecorrido
            });

            if (mesmaChapa && mesmoColetor && mesmoTipo && tempoDecorrido < 120) {
              return { ok: false, msg: 'Registro já enviado recentemente. Aguarde alguns minutos antes de reenviar.' };
            }
          } catch (dateError) {
            console.warn('[Coletores] Erro ao processar data da última entrada:', dateError);
          }
        }
      }

      // Salva o registro
      const novoRegistro = {
        'Data': agora.toISOString(),
        'Chapa': chapa,
        'Nome': nome,
        'Funcao': funcao,
        'NumeroColetor': numeroColetor,
        'TipoOperacao': tipoOperacao,
        'Situacao': situacoesTexto,
        'Supervisor': supervisor || '',
        'Coluna 2': '' // Coluna adicional conforme estrutura informada
      };

      console.log('[Coletores] Dados a serem salvos:', novoRegistro);
      
      await sheet.addRow(novoRegistro);

      console.log(`[Coletores] ✅ Registro salvo: Coletor ${numeroColetor} - ${nome}`);
      return { ok: true, msg: 'Dados salvos com sucesso!' };
    } catch (error) {
      console.error('[Coletores] Erro ao salvar registro:', error);
      return { ok: false, msg: `Erro ao salvar registro: ${error.message}` };
    }
  }

  async obterStatus() {
    try {
      console.log('[Coletores] Iniciando obtenção de status...');
      const sheet = await this.getSheet('Historico');
      const rows = await sheet.getRows();
      
      console.log(`[Coletores] Processando ${rows.length} registros do histórico`);
      
      const mapa = {};

      // Varre de trás para frente para pegar último status de cada coletor
      for (let i = rows.length - 1; i >= 0; i--) {
        const row = rows[i];
        
        const dataStr = row.get('Data') || '';
        const chapa = row.get('Chapa') || '';
        const nome = row.get('Nome') || '';
        const funcao = row.get('Funcao') || '';
        const coletor = row.get('NumeroColetor') || '';
        const tipo = row.get('TipoOperacao') || '';
        const situacao = row.get('Situacao') || '';
        const supervisor = row.get('Supervisor') || '';

        // Só adiciona se ainda não temos status para este coletor
        if (coletor && !mapa[coletor]) {
          let dataFormatada = 'Data inválida';
          
          try {
            if (dataStr) {
              const data = new Date(dataStr);
              dataFormatada = data.toLocaleDateString('pt-BR') + ' ' + data.toLocaleTimeString('pt-BR');
            }
          } catch (dateError) {
            console.warn(`[Coletores] Erro ao formatar data "${dataStr}":`, dateError);
          }

          mapa[coletor] = {
            data: dataFormatada,
            tipo: tipo || 'N/A',
            situacao: situacao || 'N/A',
            chapa: chapa || 'N/A',
            nome: nome || 'N/A',
            supervisor: supervisor || 'Sem Supervisor'
          };
        }
      }

      console.log(`[Coletores] Status processado para ${Object.keys(mapa).length} coletores`);
      return mapa;
    } catch (error) {
      console.error('[Coletores] Erro ao obter status:', error);
      throw error;
    }
  }

  async gerarResumo() {
    try {
      console.log('[Coletores] Gerando resumo...');
      const statusPorColetor = await this.obterStatus();

      let disponiveis = 0;
      let indisponiveis = 0;
      let quebrados = 0;

      for (const coletor in statusPorColetor) {
        const status = statusPorColetor[coletor];

        if (status.tipo === 'Entrega' && status.situacao === 'OK') {
          disponiveis++;
        } else if (status.tipo === 'Retirada') {
          indisponiveis++;
        }

        if (status.situacao !== 'OK') {
          quebrados++;
        }
      }

      const resumo = {
        disponiveis,
        indisponiveis,
        quebrados,
        total: Object.keys(statusPorColetor).length
      };

      console.log('[Coletores] Resumo gerado:', resumo);
      return resumo;
    } catch (error) {
      console.error('[Coletores] Erro ao gerar resumo:', error);
      throw error;
    }
  }

  // MÉTODO PARA OBTER DADOS DE PRESENÇA (da planilha principal via planilha de coletores)
  async obterDadosPresenca() {
    try {
      console.log('[Coletores] Obtendo dados de presença...');
      const sheet = await this.getSheet('Presenca');
      const rows = await sheet.getRows();
      
      console.log(`[Coletores] Processando ${rows.length} registros de presença`);
      
      const dados = [];
      for (const row of rows) {
        const data = row.get('Data') || '';
        const matricula = row.get('Matricula') || '';
        const nome = row.get('Nome') || '';
        const funcao = row.get('Função') || '';
        const supervisor = row.get('Supervisor') || '';
        const ref = row.get('REF') || '';
        const status = row.get('Status') || '';
        
        if (matricula && nome) {
          dados.push({
            data,
            matricula,
            nome,
            funcao,
            supervisor,
            ref,
            status
          });
        }
      }
      
      console.log(`[Coletores] Total de dados de presença: ${dados.length}`);
      return dados;
    } catch (error) {
      console.error('[Coletores] Erro ao obter dados de presença:', error);
      throw error;
    }
  }
}

const sheetsColetoresService = new SheetsColetoresService();
module.exports = sheetsColetoresService;
