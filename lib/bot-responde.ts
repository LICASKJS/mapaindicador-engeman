export const botMenus = {
  main: {
    text: "👋 Olá! Eu sou o assistente virtual da **Engeman**. Como posso te ajudar?\n\n📊 Escolha uma opção:",
    buttons: [
      { label: "📈 Desempenho de Fornecedores", action: "desempenho" },
      { label: "📅 Indicadores Mensais", action: "indicadores" },
      { label: "📄 Documentações Cadastrais", action: "documentos" },
      { label: "💬 Suporte e Contato", action: "suporte" },
      { label: "📋 Procedimento Engeman", action: "procedimento" },
    ],
  },
  desempenho: {
    text: "📊 **Desempenho de Fornecedores**\n\nSelecione a categoria:",
    buttons: [
      { label: "✅ Fornecedores Aprovados", action: "aprovados" },
      { label: "⚠️ Em Atenção", action: "atencao" },
      { label: "❌ Reprovados", action: "reprovados" },
      { label: "🔙 Voltar", action: "main" },
    ],
  },
  indicadores: {
    text: "📅 **Indicadores Mensais**\n\nSelecione o mês:",
    buttons: [
      { label: "📊 Ranking Mensal", action: "ranking" },
      { label: "📈 Tendência de Desempenho", action: "tendencia" },
      { label: "🔙 Voltar", action: "main" },
    ],
  },
  documentos: {
    text: "📄 **Documentações Cadastrais**\n\nSelecione uma opção:",
    buttons: [
      { label: "📋 Vencimentos de Documentos", action: "vencimentos" },
      { label: "🔙 Voltar", action: "main" },
    ],
  },
  suporte: {
    text: "💬 **Suporte e Contato**\n\n📧 Email: suporte@engeman.com.br\n📞 Telefone: (11) 3000-0000\n\n🔙 Voltar ao menu principal",
    buttons: [{ label: "🔙 Voltar", action: "main" }],
  },
  procedimento: {
    text: "📋 **Procedimento Engeman**\n\nProcedimento de avaliação de fornecedores conforme normas internas da empresa.\n\n🔙 Voltar ao menu principal",
    buttons: [{ label: "🔙 Voltar", action: "main" }],
  },
}

export async function getAprovadosFornecedores(qualidadeData: any[]) {
  // Filtra fornecedores aprovados (nota >= 7)
  const aprovados = qualidadeData
    .filter((item: any) => Number.parseFloat(item.nota) >= 7)
    .map((item: any) => item.nome_agente)
    .filter((value: string, index: number, self: string[]) => self.indexOf(value) === index)

  return `✅ **Fornecedores Aprovados**\n\n${aprovados.map((f: string) => `• ${f}`).join("\n")}`
}

export async function getAtencaoFornecedores(qualidadeData: any[]) {
  // Filtra fornecedores em atenção (nota entre 5 e 7)
  const atencao = qualidadeData
    .filter((item: any) => {
      const nota = Number.parseFloat(item.nota)
      return nota >= 5 && nota < 7
    })
    .map((item: any) => item.nome_agente)
    .filter((value: string, index: number, self: string[]) => self.indexOf(value) === index)

  return `⚠️ **Fornecedores em Atenção**\n\n${atencao.map((f: string) => `• ${f}`).join("\n")}`
}

export async function getReprovadosFornecedores(qualidadeData: any[]) {
  // Filtra fornecedores reprovados (nota < 5)
  const reprovados = qualidadeData
    .filter((item: any) => Number.parseFloat(item.nota) < 5)
    .map((item: any) => item.nome_agente)
    .filter((value: string, index: number, self: string[]) => self.indexOf(value) === index)

  return `❌ **Fornecedores Reprovados**\n\n${reprovados.map((f: string) => `• ${f}`).join("\n")}`
}
