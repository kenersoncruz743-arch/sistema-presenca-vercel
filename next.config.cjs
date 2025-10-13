// next.config.js - Atualizado para incluir rota de coletores
/** @type {import('next').NextConfig} */
const nextConfig = {
  reactStrictMode: false,
  
  async rewrites() {
    return [
      // Rota para login (página inicial)
      {
        source: '/',
        destination: '/index.html',
      },
      // Rota para menu
      {
        source: '/menu',
        destination: '/menu.html',
      },
      // Rota para painel de presença
      {
        source: '/painel',
        destination: '/painel.html',
      },
      // NOVA ROTA: Rota para controle de coletores
      {
        source: '/coletores',
        destination: '/coletores.html',
      },
      {
        source: '/qlp',
        destination: '/qlp.html',
      }
    ];
  },
}

module.exports = nextConfig;
