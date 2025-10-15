// next.config.js - Atualizado com rota de Mapa de Carga
/** @type {import('next').NextConfig} */
const nextConfig = {
  reactStrictMode: false,
  
  async rewrites() {
    return [
      {
        source: '/',
        destination: '/index.html',
      },
      {
        source: '/ferramentas',
        destination: '/ferramentas.html',
      },
      
      {
        source: '/menu',
        destination: '/menu.html',
      },
      {
        source: '/painel',
        destination: '/painel.html',
      },
      {
        source: '/coletores',
        destination: '/coletores.html',
      },
      {
        source: '/qlp',
        destination: '/qlp.html',
      },
      // NOVA ROTA: Mapa de Carga
      {
        source: '/mapacarga',
        destination: '/mapacarga.html',
      },
      
      {
        source: '/producao',
        destination: '/producao.html',
      }
      
    ];
  },
}

module.exports = nextConfig;
