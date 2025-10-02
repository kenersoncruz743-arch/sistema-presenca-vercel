// pages/index.js
// Este componente é apenas um placeholder para satisfazer o Next.js.

// 1. Importa a função useEffect (para rodar código após a renderização)
import { useEffect } from 'react';

// 2. Define o componente (obrigatório pelo Next.js)
export default function Home() {
  
  // 3. Redireciona para o arquivo estático no /public/
  useEffect(() => {
    // A rota '/' no Next.js é usada para a página inicial.
    // Usamos window.location para forçar o navegador a ir para o 
    // seu index.html, que é servido via rewrites.
    window.location.href = '/index.html'; 
  }, []);

  // 4. Retorna um JSX mínimo (obrigatório pelo Next.js)
  return (
    <div style={{ padding: '50px', textAlign: 'center' }}>
      Redirecionando para o Login...
    </div>
  );
}
