// api/producao/salvar.js - Salva dados na aba Produzido
const sheetsService = require('../../lib/sheets');

module.exports = async function handler(req, res) {
  res.setHeader('Access-Control-Allow-Origin', '*');
  res.setHeader('Access-Control-Allow-Methods', 'POST, OPTIONS');
  res.setHeader('Access-Control-Allow-Headers', 'Content-Type');

  if (req.method === 'OPTIONS') {
    return res.status(200).end();
  }

  if (req.method !== 'POST') {
    return res.status(405).json({ 
      ok: false, 
      msg: 'Método não permitido' 
    });
  }

  try {
    const { dados, dataReferencia, horasTrabalhadas } = req.body;
    
    if (!dados || !Array.isArray(dados)) {
      return res.status(400).json({
        ok: false,
        msg: 'Dados inválidos'
      });
    }
    
    console.log('[PRODUCAO/SALVAR] Iniciando salvamento...');
    console.log(`[PRODUCAO/SALVAR] ${dados.length} registros para salvar`);
    
    const doc = await sheetsService.init();
    console.log(`[PRODUCAO/SALVAR] Conectado: ${doc.title}`);
    
    // Busca ou cria aba Produzido
    let sheetProduzido = doc.sheetsByTitle['Produzido'];
    
    if (!sheetProduzido) {
      console.log('[PRODUCAO/SALVAR] Criando aba Produzido...');
      sheetProduzido = await doc.addSheet({
        title: 'Produzido',
        headerValues: [
          'Data',
          'Seção',
          'Supervisor',
          'Turno',
          'QL',
          'Capacidade',
          'Presente',
          'Desvio Função',
          'Vagas Pendentes',
          'Afastado',
          'Subtotal Ausente',
          'Horas Trabalhadas',
          'Produtividade/Hora',
          'Realizado',
          'Produzido'
        ]
      });
    }
    
    const dataRef = dataReferencia || new Date().toLocaleDateString('pt-BR');
    const horas = horasTrabalhadas || 8;
    
    console.log(`[PRODUCAO/SALVAR] Data: ${dataRef}, Horas: ${horas}`);
    
    // Carrega registros existentes
    const rowsExistentes = await sheetProduzido.getRows();
    console.log(`[PRODUCAO/SALVAR] ${rowsExistentes.length} registros existentes`);
    
    let totalNovos = 0;
    let totalAtualizados = 0;
    
    for (const item of dados) {
      // Busca registro existente (mesma data, seção, turno e supervisor)
      let registroExistente = null;
      
      for (const row of rowsExistentes) {
        const rowData = String(row.get('Data') || '').trim();
        const rowSecao = String(row.get('Seção') || '').trim();
        const rowTurno = String(row.get('Turno') || '').trim();
        const rowSupervisor = String(row.get('Supervisor') || '').trim();
        
        if (rowData === dataRef && 
            rowSecao === item.secao && 
            rowTurno === item.turno && 
            rowSupervisor === item.supervisor) {
          registroExistente = row;
          break;
        }
      }
      
      if (registroExistente) {
        // Atualiza registro existente
        console.log(`[PRODUCAO/SALVAR] Atualizando: ${item.secao} - ${item.turno}`);
        
        registroExistente.set('QL', item.ql);
        registroExistente.set('Capacidade', item.capacidade);
        registroExistente.set('Presente', item.presente);
        registroExistente.set('Desvio Função', item.desvioFuncao);
        registroExistente.set('Vagas Pendentes', item.vagasPendentes);
        registroExistente.set('Afastado', item.afastado);
        registroExistente.set('Subtotal Ausente', item.subtotalAusente);
        registroExistente.set('Horas Trabalhadas', horas);
        registroExistente.set('Produtividade/Hora', item.produtividadeHora);
        registroExistente.set('Realizado', item.realizado);
        registroExistente.set('Produzido', item.produzido);
        
        await registroExistente.save();
        totalAtualizados++;
        
      } else {
        // Insere novo registro
        console.log(`[PRODUCAO/SALVAR] Inserindo: ${item.secao} - ${item.turno}`);
        
        await sheetProduzido.addRow({
          'Data': dataRef,
          'Seção': item.secao,
          'Supervisor': item.supervisor,
          'Turno': item.turno,
          'QL': item.ql,
          'Capacidade': item.capacidade,
          'Presente': item.presente,
          'Desvio Função': item.desvioFuncao,
          'Vagas Pendentes': item.vagasPendentes,
          'Afastado': item.afastado,
          'Subtotal Ausente': item.subtotalAusente,
          'Horas Trabalhadas': horas,
          'Produtividade/Hora': item.produtividadeHora,
          'Realizado': item.realizado,
          'Produzido': item.produzido
        });
        
        totalNovos++;
      }
    }
    
    console.log('[PRODUCAO/SALVAR] Salvamento concluído!');
    console.log(`[PRODUCAO/SALVAR] ${totalNovos} novos, ${totalAtualizados} atualizados`);
    
    return res.status(200).json({
      ok: true,
      msg: `Dados salvos com sucesso! ${totalNovos} novos registros, ${totalAtualizados} atualizados.`,
      totais: {
        novos: totalNovos,
        atualizados: totalAtualizados,
        total: totalNovos + totalAtualizados
      }
    });
    
  } catch (error) {
    console.error('[PRODUCAO/SALVAR] Erro:', error);
    return res.status(500).json({
      ok: false,
      msg: 'Erro ao salvar dados',
      details: error.message,
      stack: process.env.NODE_ENV === 'development' ? error.stack : undefined
    });
  }
};
