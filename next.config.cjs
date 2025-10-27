// next.config.js - Configuração completa de rotas
/** @type {import('next').NextConfig} */
const nextConfig = {
  reactStrictMode: false,
  
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
        source: '/coletores',
        destination: '/coletores.html',
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
