import { GoogleSpreadsheet } from 'google-spreadsheet';
import { JWT } from 'google-auth-library';

export default async function handler(req, res) {
  try {
    console.log('🔍 Testando conexão com Google Sheets...');
    
    // Verifica se as variáveis existem
    const hasId = !!process.env.GOOGLE_SHEETS_ID;
    const hasEmail = !!process.env.GOOGLE_SHEETS_CLIENT_EMAIL;
    const hasKey = !!process.env.GOOGLE_SHEETS_PRIVATE_KEY;
    
    console.log('Variáveis configuradas:', { hasId, hasEmail, hasKey });
    
    if (!hasId || !hasEmail || !hasKey) {
      return res.status(500).json({
        error: 'Variáveis de ambiente não configuradas',
        hasId,
        hasEmail,
        hasKey
      });
    }
    
    // Tenta conectar
    const serviceAccountAuth = new JWT({
      email: process.env.GOOGLE_SHEETS_CLIENT_EMAIL,
      key: process.env.GOOGLE_SHEETS_PRIVATE_KEY.replace(/\\n/g, '\n'),
      scopes: ['https://www.googleapis.com/auth/spreadsheets'],
    });

    const doc = new GoogleSpreadsheet(process.env.GOOGLE_SHEETS_ID, serviceAccountAuth);
    await doc.loadInfo();
    
    return res.status(200).json({
      success: true,
      message: 'Conexão bem-sucedida!',
      planilha: doc.title,
      abas: doc.sheetsByIndex.map(s => s.title)
    });
    
  } catch (error) {
    console.error('❌ Erro na conexão:', error);
    return res.status(500).json({
      error: error.message,
      stack: error.stack,
      name: error.name
    });
  }
}
