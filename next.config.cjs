// next.config.js - Configuração completa de rotas e CORS
/** @type {import('next').NextConfig} */
const nextConfig = {
  reactStrictMode: false,
  
  // Headers globais para CORS
  async headers() {
    return [
      {
        source: '/api/:path*',
        headers: [
          { key: 'Access-Control-Allow-Origin', value: '*' },
          { key: 'Access-Control-Allow-Methods', value: 'GET, POST, PUT, DELETE, OPTIONS' },
          { key: 'Access-Control-Allow-Headers', value: 'Content-Type, Authorization, X-Requested-With, Accept' },
          { key: 'Access-Control-Allow-Credentials', value: 'true' },
        ],
      },
    ];
  },
  
  async rewrites() {
    return [
      // Página inicial
      {
        source: '/',
        destination: '/index.html',
      },
      
      // Menu e ferramentas
      {
        source: '/menu',
        destination: '/menu.html',
      },
      {
        source: '/ferramentas',
        destination: '/ferramentas.html',
      },
      
      // Ferramentas principais
      {
        source: '/painel',
        destination: '/painel.html',
      },
      {
        source: '/resumo-equipamentos',
        destination: '/resumo-equipamentos.html',
      },
      {
        source: '/controle-coletores',
        destination: '/controle-coletores.html',
      },
      
      // QLP e análises
      {
        source: '/qlp',
        destination: '/qlp.html',
      },
      
      // Produção
      {
        source: '/producao',
        destination: '/producao.html',
      },
      {
        source: '/resumo-base',
        destination: '/resumo-base.html',
      },
      
      // Logística
      {
        source: '/mapacarga',
        destination: '/mapacarga.html',
      },
      {
        source: '/alocacaobox',
        destination: '/alocacaobox.html',
      }
    ];
  },
}

module.exports = nextConfig;
