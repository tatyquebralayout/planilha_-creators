/**
 * CONSTANTES GLOBAIS
 * Definições e configurações usadas em todo o sistema
 */

// Constantes de Status e Estados
const STATUS_OPTIONS = [
  "Não Contatado",
  "Contatado", 
  "Aguardando Resposta",
  "Reunião Agendada",
  "Confirmado",
  "Não Interessado",
  "Parceria Fechada",
  "Parceria Ativa",
  "Em Renovação",
  "Encerrado"
];

const STATUS_CONTEUDO = [
  "Planejado",
  "Em Produção",
  "Aprovado",
  "Publicado",
  "Em Análise",
  "Concluído"
];

const STATUS_QUALIFICACAO = [
  "Excelente (20-25)",
  "Bom (15-19)",
  "Regular (10-14)",
  "Baixo Potencial (<10)"
];

// Constantes de Categorização
const PRIORIDADE_OPTIONS = [
  "Nível 1 (Alta)",
  "Nível 2 (Média)", 
  "Nível 3 (Baixa)",
  "Lista de Espera"
];

const ETAPAS_JORNADA = [
  "Identificação do Problema",
  "Exploração da Solução",
  "Construção de Requisitos",
  "Seleção",
  "Compra/Conversão",
  "Retenção",
  "Advocacia"
];

const NIVEL_PERSONALIZACAO = [
  "Básico",
  "Intermediário",
  "Avançado",
  "Premium"
];

const SEGMENTO_OPTIONS = [
  "Macro",
  "Estratégico",
  "Crescimento",
  "Oportunidade"
];

// Constantes de Conteúdo e Plataformas
const FORMATO_CONTEUDO = [
  "Post Feed",
  "Story",
  "Reels/TikTok",
  "Live",
  "IGTV/YouTube",
  "Podcast",
  "Blog",
  "Newsletter",
  "Webinar",
  "Evento Presencial"
];

const PLATAFORMAS = [
  "Instagram",
  "TikTok",
  "YouTube",
  "Twitter",
  "LinkedIn",
  "Twitch",
  "Facebook",
  "Pinterest",
  "Podcast",
  "Blog",
  "Newsletter",
  "Outra"
];

// Constantes de Tempo e Períodos
const CHECKPOINT_OPTIONS = [
  "1º Trim 2023",
  "2º Trim 2023",
  "3º Trim 2023",
  "4º Trim 2023",
  "1º Trim 2024",
  "2º Trim 2024",
  "3º Trim 2024",
  "4º Trim 2024"
];

/**
 * FUNÇÕES DE INICIALIZAÇÃO E CONFIGURAÇÃO
 */

/**
 * Inicializa o menu personalizado na planilha
 */
function onOpen() {
  var ui = SpreadsheetApp.getUi();
  
  ui.createMenu('Gestão de Creators')
    .addItem('Criar Planilha Completa', 'criarPlanilhaCompleta')
    .addSeparator()
    .addItem('Adicionar Creator Exemplo', 'criarCreatorExemplo')
    .addItem('Atualizar Estatísticas', 'atualizarEstatisticas')
    .addSeparator()
    .addItem('Exportar Calendário', 'exportarCalendario')
    .addItem('Sincronizar Calendário', 'sincronizarCalendario')
    .addSeparator()
    .addItem('Gerar Relatório de Desempenho', 'gerarRelatorioDesempenho')
    .addItem('Gerar Análise e Insights', 'gerarAnaliseInsights')
    .addSeparator()
    .addItem('Verificar Tarefas Pendentes', 'verificarTarefasPendentes')
    .addItem('Limpar Dados Antigos', 'limparDadosAntigos')
    .addSeparator()
    .addItem('Configurar Gatilhos Automáticos', 'configurarGatilhosAutomaticos')
    .addItem('Gerar Backup', 'gerarBackup')
    .addToUi();
}

/**
 * Função principal para criar/redefinir todas as planilhas
 */
function criarPlanilhaCompleta() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();

  var guias = [
    // Página de Instruções e Apresentação
    {nome: "Página de Instruções", colunas: ["Instruções de Uso da Planilha"]},
    {nome: "Sumário Executivo", colunas: ["Resumo do Programa de Creators", "Indicadores Principais", "Período Atual", "Status Geral"]},
    {nome: "Jornada do Seguidor", colunas: ["Mapeamento da Jornada de Consumo"]},
    
    // Gestão de Creators (núcleo)
    {nome: "Guia Principal", colunas: ["ID", "Nome", "Categoria", "Nicho Específico", "Plataforma Principal", "Seguidores", "Engajamento (%)", "Email", "WhatsApp", "Link Perfil", "Data Primeiro Contato", "Status Contato", "Próxima Ação", "Data Reunião", "Horário Reunião", "Responsável Interno", "Prioridade", "Etapa Jornada Compra", "Nível Personalização", "Observações"]},
    {nome: "Calendário de Reuniões", colunas: ["Data", "Horário", "Nome Creator", "Plataforma (Zoom/Meet)", "Link da Reunião", "Status da Reunião"]},
    {nome: "Histórico de Campanhas", colunas: ["ID Creator", "Nome Creator", "Data Campanha", "Tipo Campanha", "Resultado", "Métricas de Desempenho", "ROI Estimado", "Feedback Creator"]},
    
    // Análise e Qualificação
    {nome: "Qualificação dos Creators", colunas: ["Nome Creator", "Potencial Alcance (1-5)", "Relevância para Audiência (1-5)", "Compatibilidade com Marca (1-5)", "Qualidade de Conteúdo (1-5)", "Engajamento da Audiência (1-5)", "Histórico de Conversão (1-5)", "Pontuação Total", "Status Qualificação", "Segmento", "Nível de Prioridade"]},
    {nome: "Perfil Ideal de Creator", colunas: ["Critério", "Perfil Ideal", "Creator 1", "Creator 2", "Creator 3", "Creator 4", "Creator 5", "Creator 6", "Creator 7", "Creator 8"]},
    {nome: "Perfil Detalhado do Creator", colunas: ["Nome Creator", "Papel na Decisão", "Objetivos Pessoais", "Riscos ou Objeções", "Estilo de Comunicação", "Influenciadores", "Necessidades Específicas", "Histórico de Relacionamento"]},
    {nome: "Análise de Interesses", colunas: ["Creators Envolvidos", "Interesses Compartilhados", "Pontos de Divergência", "Potencial de Colaboração", "Ações Recomendadas"]},
    
    // Planejamento Estratégico
    {nome: "Alinhamento Estratégico", colunas: ["Objetivo de Negócio", "Estratégia com Creators", "Indicadores", "Métricas de Sucesso", "Tática", "Responsável", "Status"]},
    {nome: "Alocação de Recursos", colunas: ["Categoria", "Orçamento Estimado", "Orçamento Realizado", "Variação", "Observações"]},
    {nome: "Manual de Engajamento", colunas: ["Etapa da Jornada", "Ações de Marketing", "Ações de Vendas", "Canais Específicos", "Responsável pela Ação"]},
    
    // Planejamento de Conteúdo
    {nome: "Plano de Conteúdo", colunas: ["Tipo de Conteúdo", "Nível de Personalização", "Detalhes da Personalização", "Personalização por Segmento", "Personalização por Creator", "Recursos Necessários"]},
    {nome: "Calendário Editorial", colunas: ["Semana", "Creator", "Tema", "Formato", "Plataforma", "Data Publicação", "Status", "Link do Conteúdo", "Métricas", "Observações"]},
    {nome: "Modelo de Mensagens", colunas: ["Tipo Mensagem", "Texto Modelo", "Tempo Médio Resposta (dias)", "Tipo de Resposta", "Etapa da Prospecção", "Responsável", "Variantes para Testes"]},
    
    // Métricas e Análises
    {nome: "Métricas de Desempenho", colunas: ["Categoria", "Tipo", "Métrica", "Checkpoint 1", "Checkpoint 2", "Checkpoint 3", "Checkpoint 4", "Checkpoint 5", "Meta", "Realizado", "Status"]},
    {nome: "Estatísticas Gerais", colunas: ["Total Creators Captados", "Total Reuniões Agendadas", "Total Confirmados", "Média Engajamento (%)", "Taxa de Conversão (%)", "ROI Global", "Checkpoint", "Meta Checkpoint", "Realizado"]},
    {nome: "Dashboard", colunas: ["Dashboard Visual Avançado"]},
    
    // Acompanhamento e Implementação
    {nome: "Cronograma Semanal", colunas: ["Semana", "Creator", "Atividades Críticas", "Responsável", "Status", "Observações", "Data Conclusão", "Prioridade"]}
  ];

  // Processamento em lotes para não sobrecarregar a execução
  var batchSize = 5;
  var batchCount = Math.ceil(guias.length / batchSize);
  
  for (var batch = 0; batch < batchCount; batch++) {
    var startIdx = batch * batchSize;
    var endIdx = Math.min(startIdx + batchSize, guias.length);
    
    for (var i = startIdx; i < endIdx; i++) {
      var guia = guias[i];
      var sheet = ss.getSheetByName(guia.nome) || ss.insertSheet(guia.nome);
      sheet.clear();
      
      // Adiciona cabecalho e formata
      sheet.appendRow(guia.colunas);
      var headerRange = sheet.getRange(1, 1, 1, guia.colunas.length);
      headerRange.setBackground("#4285F4");
      headerRange.setFontColor("white");
      headerRange.setFontWeight("bold");
      
      // Congela a primeira linha
      sheet.setFrozenRows(1);
      
      // Ajusta largura das colunas automaticamente
      sheet.autoResizeColumns(1, guia.colunas.length);
      
      // Aplica formatação específica para cada guia
      aplicarFormatacaoEspecifica(sheet, guia.nome);
    }
    
    // Pausa entre lotes para não sobrecarregar
    Utilities.sleep(500);
  }
  
  // Aplica regras de validação
  aplicarRegraValidacao();
  
  // Inicializa páginas especiais
  inicializarPaginaInstrucoes();
  inicializarSumarioExecutivo();
  inicializarPlaybookEngajamento();
  inicializarAnaliseInteresses();
  
  // Configura o dashboard
  configurarDashboard();

  // Atualiza estatísticas
  atualizarEstatisticas();

  // Exibe mensagem de confirmação
  SpreadsheetApp.getUi().alert("Planilha criada com sucesso!\n\nUtilize o menu 'Gestão de Creators' para acessar todas as funcionalidades.");
}

/**
 * Inicializa a página de instruções
 */
function inicializarPaginaInstrucoes() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("Página de Instruções");
  
  if (!sheet) {
    return;
  }
  
  sheet.clear();
  
  // Título e cabeçalhos
  sheet.getRange("A1:F1").merge();
  sheet.getRange("A1").setValue("PLANILHA DE GESTÃO DE CREATORS");
  sheet.getRange("A1").setFontSize(18).setFontWeight("bold").setHorizontalAlignment("center");
  sheet.getRange("A1:F1").setBackground("#4285F4").setFontColor("white");
  
  sheet.getRange("A2:F2").merge();
  sheet.getRange("A2").setValue("Instruções de Uso");
  sheet.getRange("A2").setFontSize(14).setFontWeight("bold").setHorizontalAlignment("center");
  
  // Dimensionar colunas
  sheet.setColumnWidth(1, 30);
  sheet.setColumnWidth(2, 200);
  sheet.setColumnWidth(3, 350);
  sheet.setColumnWidth(4, 200);
  
  // Visão geral da planilha
  sheet.getRange("B4").setValue("VISÃO GERAL");
  sheet.getRange("B4").setFontWeight("bold");
  sheet.getRange("C4").setValue("Esta planilha foi desenvolvida para gerenciar o processo completo de prospecção, qualificação e gestão de parcerias com creators. Siga as instruções abaixo para utilizar cada seção de forma eficiente.");
  
  // Estrutura das instruções
  var instrucoes = [
    ["INÍCIO RÁPIDO", "Para começar a usar a planilha:", "1. Vá para a guia 'Guia Principal' e comece a adicionar creators\n2. Use as validações de dados para manter a consistência\n3. Acesse o Dashboard para visualizar métricas gerais\n4. Utilize o menu personalizado em 'Gestão de Creators' para acessar funcionalidades adicionais"],
    
    ["GUIAS PRINCIPAIS", "Guia Principal", "Registro central de todos os creators, com informações de contato, métricas e status de prospecção"],
    ["", "Qualificação dos Creators", "Pontuação e segmentação de creators com base em critérios de qualificação que ajudam a priorizar esforços"],
    ["", "Calendário de Reuniões", "Visualização e gestão de reuniões agendadas com creators"],
    ["", "Dashboard", "Visualização consolidada das métricas e KPIs principais do programa de creators"],
    
    ["PLANEJAMENTO", "Alinhamento Estratégico", "Defina objetivos estratégicos e alinhe suas ações com creators aos objetivos de negócio"],
    ["", "ICP de Creators", "Defina o perfil ideal de creators para sua marca e compare prospects com este perfil"],
    ["", "Plano de Conteúdo", "Planeje o conteúdo a ser produzido com creators ao longo da jornada de compra"],
    ["", "Alocação de Recursos", "Gerencie e acompanhe o orçamento dedicado às ações com creators"],
    
    ["EXECUÇÃO", "Histórico de Campanhas", "Registro de campanhas anteriores com creators, resultados e aprendizados"],
    ["", "Calendário Editorial", "Planejamento detalhado de publicações e criação de conteúdo"],
    ["", "Playbook de Engajamento", "Estratégias específicas para cada etapa da jornada do creator"],
    
    ["ANÁLISE", "Métricas de Desempenho", "Acompanhamento detalhado de todas as métricas relevantes do programa"],
    ["", "Estatísticas Gerais", "Visão consolidada das estatísticas e métricas"],
    ["", "Análise Competitiva", "Compare seu relacionamento com creators em relação à concorrência"],
    
    ["MENU PERSONALIZADO", "O menu 'Gestão de Creators'", "Oferece acesso rápido a funcionalidades como:\n- Criar/Redefinir Planilha: Redefine todas as guias com estrutura padrão\n- Atualizar Estatísticas: Atualiza todas as métricas automaticamente\n- Gerar Relatório: Cria um relatório resumido das atividades\n- Exportar Calendário: Exporta reuniões para o Google Calendar"]
  ];
  
  // Inserir as instruções
  var startRow = 6;
  var currentSection = "";
  
  for (var i = 0; i < instrucoes.length; i++) {
    var row = startRow + i;
    
    // Se for uma nova seção
    if (instrucoes[i][0] !== "") {
      currentSection = instrucoes[i][0];
      sheet.getRange(row, 2).setValue(currentSection);
      sheet.getRange(row, 2).setFontWeight("bold").setBackground("#e3f2fd");
      row++;
    }
    
    // Inserir guia/funcionalidade e instrução
    sheet.getRange(row, 3).setValue(instrucoes[i][1]);
    sheet.getRange(row, 3).setFontWeight("bold");
    sheet.getRange(row, 4).setValue(instrucoes[i][2]);
    sheet.getRange(row, 4).setWrap(true);
  }
  
  // Adiciona dicas de uso
  var lastRow = startRow + instrucoes.length + 3;
  sheet.getRange(lastRow, 2).setValue("DICAS PARA MELHORES RESULTADOS");
  sheet.getRange(lastRow, 2).setFontWeight("bold");
  
  var dicas = [
    "• Mantenha os dados atualizados regularmente para garantir relatórios precisos",
    "• Use a função de qualificação para focar nos creators mais promissores",
    "• Documente todas as interações na guia 'Histórico de Campanhas'",
    "• Utilize as regras de validação de dados para manter a consistência",
    "• Agende reuniões regulares de revisão usando os dados da planilha"
  ];
  
  for (var i = 0; i < dicas.length; i++) {
    sheet.getRange(lastRow + 1 + i, 3, 1, 2).merge();
    sheet.getRange(lastRow + 1 + i, 3).setValue(dicas[i]);
  }
  
  // Adiciona rodapé
  var footerRow = lastRow + dicas.length + 3;
  sheet.getRange(footerRow, 2, 1, 3).merge();
  sheet.getRange(footerRow, 2).setValue("Para suporte ou dúvidas, entre em contato com o responsável pela planilha");
  sheet.getRange(footerRow, 2).setHorizontalAlignment("center").setFontStyle("italic");
}

/**
 * Inicializa a página de resumo executivo
 */
function inicializarSumarioExecutivo() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("Sumário Executivo");
  
  if (!sheet) {
    return;
  }
  
  sheet.clear();
  
  // Configuração básica da página
  sheet.setColumnWidth(1, 30);
  sheet.setColumnWidth(2, 200);
  sheet.setColumnWidth(3, 200);
  sheet.setColumnWidth(4, 200);
  sheet.setColumnWidth(5, 200);
  
  // Cabeçalho
  sheet.getRange("B1:E1").merge();
  sheet.getRange("B1").setValue("SUMÁRIO EXECUTIVO - PROGRAMA DE MARKETING COM CREATORS");
  sheet.getRange("B1").setFontSize(16).setFontWeight("bold").setHorizontalAlignment("center");
  sheet.getRange("B1:E1").setBackground("#4285F4").setFontColor("white");
  
  // Data do relatório
  var today = new Date();
  sheet.getRange("B2:E2").merge();
  sheet.getRange("B2").setValue("Período: " + Utilities.formatDate(today, "GMT-3", "MMMM yyyy"));
  sheet.getRange("B2").setHorizontalAlignment("center").setFontStyle("italic");
  
  // Visão geral
  sheet.getRange("B4").setValue("VISÃO GERAL DO PROGRAMA");
  sheet.getRange("B4").setFontWeight("bold");
  sheet.getRange("B4:E4").setBackground("#e3f2fd");
  
  sheet.getRange("B5:E8").merge();
  sheet.getRange("B5").setValue("[Insira aqui um resumo executivo do programa de marketing com creators, destacando os principais objetivos, estratégias e resultados esperados]");
  sheet.getRange("B5").setFontStyle("italic").setFontColor("#666666");
  
  // Indicadores chave
  sheet.getRange("B10").setValue("INDICADORES CHAVE DE DESEMPENHO (KPIs)");
  sheet.getRange("B10").setFontWeight("bold");
  sheet.getRange("B10:E10").setBackground("#e3f2fd");
  
  var kpis = [
    ["Métrica", "Meta", "Realizado", "Status"],
    ["Total de Creators Ativos", "=50", "=COUNTIF('Guia Principal'!L2:L, \"Parceria Ativa\")", "=IF(C12/B12>=1,\"Atingido\",\"Em Progresso\")"],
    ["Novas Parcerias (mês atual)", "=10", "=COUNTIFS('Guia Principal'!L2:L, \"Parceria Fechada\", 'Guia Principal'!K2:K, \">=\" & DATE(YEAR(TODAY()), MONTH(TODAY()), 1))", "=IF(C13/B13>=1,\"Atingido\",\"Em Progresso\")"],
    ["Taxa de Conversão", "=20%", "=IFERROR(COUNTIF('Guia Principal'!L2:L, \"Parceria Fechada\")/COUNTA('Guia Principal'!B2:B),0)", "=IF(C14/B14>=1,\"Atingido\",\"Em Progresso\")"],
    ["Engajamento Médio", "=5%", "=AVERAGE('Guia Principal'!G2:G)", "=IF(C15/B15>=1,\"Atingido\",\"Em Progresso\")"],
    ["ROI do Programa", "=200%", "=IFERROR(AVERAGE('Histórico de Campanhas'!G2:G)*100,0)", "=IF(C16/B16>=1,\"Atingido\",\"Em Progresso\")"]
  ];
  
  sheet.getRange(11, 2, kpis.length, kpis[0].length).setValues(kpis);
  sheet.getRange(11, 2, 1, kpis[0].length).setFontWeight("bold").setBackground("#f5f5f5");
  
  // Formatação para métricas
  sheet.getRange("C12:D16").setHorizontalAlignment("center");
  sheet.getRange("B14:C14").setNumberFormat("0.00%");
  sheet.getRange("B15:C15").setNumberFormat("0.00%");
  sheet.getRange("B16:C16").setNumberFormat("0.00%");
  
  // Formatação condicional para status
  var range = sheet.getRange("E12:E16");
  var rule = SpreadsheetApp.newConditionalFormatRule()
    .whenTextEqualTo("Atingido")
    .setBackground("#c8e6c9")
    .build();
  var rules = [rule];
  range.setConditionalFormatRules(rules);
  
  // Principais resultados e insights
  sheet.getRange("B18").setValue("PRINCIPAIS RESULTADOS E INSIGHTS");
  sheet.getRange("B18").setFontWeight("bold");
  sheet.getRange("B18:E18").setBackground("#e3f2fd");
  
  sheet.getRange("B19:E22").merge();
  sheet.getRange("B19").setValue("[Insira aqui os principais resultados e insights obtidos no período, incluindo análises relevantes e recomendações estratégicas]");
  sheet.getRange("B19").setFontStyle("italic").setFontColor("#666666");
  
  // Próximos passos
  sheet.getRange("B24").setValue("PRÓXIMOS PASSOS E AÇÕES PRIORITÁRIAS");
  sheet.getRange("B24").setFontWeight("bold");
  sheet.getRange("B24:E24").setBackground("#e3f2fd");
  
  var acoes = [
    ["Ação", "Responsável", "Prazo", "Status"],
    ["[Ação 1]", "", "", "Não Iniciado"],
    ["[Ação 2]", "", "", "Não Iniciado"],
    ["[Ação 3]", "", "", "Não Iniciado"]
  ];
  
  sheet.getRange(25, 2, acoes.length, acoes[0].length).setValues(acoes);
  sheet.getRange(25, 2, 1, acoes[0].length).setFontWeight("bold").setBackground("#f5f5f5");
  
  // Resumo do orçamento
  sheet.getRange("B30").setValue("RESUMO DO ORÇAMENTO");
  sheet.getRange("B30").setFontWeight("bold");
  sheet.getRange("B30:E30").setBackground("#e3f2fd");
  
  var orcamento = [
    ["Categoria", "Orçamento", "Realizado", "% Utilizado"],
    ["Total do Programa", "=IFERROR(SUM('Alocação de Recursos'!B2:B9),0)", "=IFERROR(SUM('Alocação de Recursos'!C2:C9),0)", "=IFERROR(C32/B32,0)"]
  ];
  
  sheet.getRange(31, 2, orcamento.length, orcamento[0].length).setValues(orcamento);
  sheet.getRange(31, 2, 1, orcamento[0].length).setFontWeight("bold").setBackground("#f5f5f5");
  
  // Formatação para orçamento
  sheet.getRange("B32:C32").setNumberFormat("R$ #,##0.00");
  sheet.getRange("D32").setNumberFormat("0.00%");
  
  // Adiciona botão para atualizar
  sheet.getRange("E2").setValue("Atualizar Sumário");
  sheet.getRange("E2").setFontWeight("bold").setBackground("#e3f2fd");
}

/**
 * Configura o painel de controle com gráficos e métricas avançadas
 */
function configurarDashboard() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var dashSheet = ss.getSheetByName("Painel de Controle");
  
  // Limpa o painel existente
  dashSheet.clear();
  
  // Configuração básica da planilha
  dashSheet.setColumnWidth(1, 200);
  dashSheet.setColumnWidth(2, 150);
  dashSheet.setColumnWidth(3, 150);
  dashSheet.setColumnWidth(4, 150);
  dashSheet.setColumnWidth(5, 150);
  
  // Adiciona título e subtítulo
  dashSheet.getRange("A1:E1").merge();
  dashSheet.getRange("A1").setValue("PAINEL DE CONTROLE - GESTÃO DE CREATORS");
  dashSheet.getRange("A1").setFontSize(16).setFontWeight("bold").setHorizontalAlignment("center");
  dashSheet.getRange("A1:E1").setBackground("#4285F4").setFontColor("white");
  
  dashSheet.getRange("A2:E2").merge();
  dashSheet.getRange("A2").setValue("Visão Consolidada do Programa de Marketing de Creators");
  dashSheet.getRange("A2").setFontStyle("italic").setHorizontalAlignment("center");
  
  // Adiciona data da atualização
  dashSheet.getRange("A3:E3").merge();
  var today = new Date();
  dashSheet.getRange("A3").setValue("Atualizado em: " + Utilities.formatDate(today, "GMT-3", "dd/MM/yyyy"));
  dashSheet.getRange("A3").setHorizontalAlignment("center").setFontStyle("italic");
  
  // ==== INDICADORES PRINCIPAIS ====
  dashSheet.getRange("A5:E5").merge();
  dashSheet.getRange("A5").setValue("INDICADORES PRINCIPAIS");
  dashSheet.getRange("A5").setFontWeight("bold").setHorizontalAlignment("center");
  dashSheet.getRange("A5:E5").setBackground("#e3f2fd");
  
  // Layout de 2x3 para Indicadores
  var kpiLayout = [
    ["Total de Creators", "Parcerias Ativas", "ROI Médio"],
    ["=COUNTA('Guia Principal'!B2:B)", "=COUNTIF('Guia Principal'!L2:L, \"Parceria Ativa\")", "=IFERROR(AVERAGEIF('Histórico de Campanhas'!F2:F, \"<>\", 'Histórico de Campanhas'!G2:G), 0)"]
  ];
  
  dashSheet.getRange("A6:C6").setValues([kpiLayout[0]]);
  dashSheet.getRange("A7:C7").setFormulas([kpiLayout[1]]);
  
  var kpiLayout2 = [
    ["Taxa de Conversão", "Crescimento MoM", "Média Engajamento"],
    ["=COUNTIF('Guia Principal'!L2:L, \"Parceria Fechada\")/COUNTA('Guia Principal'!B2:B)", "", "=AVERAGE('Guia Principal'!G2:G)"]
  ];
  
  dashSheet.getRange("A8:C8").setValues([kpiLayout2[0]]);
  dashSheet.getRange("A9:C9").setFormulas([kpiLayout2[1]]);
  
  // Formatação dos KPIs
  dashSheet.getRange("A6:C6").setFontWeight("bold");
  dashSheet.getRange("A7:C7").setFontSize(14);
  dashSheet.getRange("A8:C8").setFontWeight("bold");
  dashSheet.getRange("A9:C9").setFontSize(14);
  
  // Formatos de números
  dashSheet.getRange("A7").setNumberFormat("#,##0");
  dashSheet.getRange("B7").setNumberFormat("#,##0");
  dashSheet.getRange("C7").setNumberFormat("#,##0.0x");
  dashSheet.getRange("A9").setNumberFormat("0.00%");
  dashSheet.getRange("B9").setNumberFormat("0.00%");
  dashSheet.getRange("C9").setNumberFormat("0.00%");
  
  // ==== SEÇÃO 1: FUNIL DE CREATORS ====
  dashSheet.getRange("A11:E11").merge();
  dashSheet.getRange("A11").setValue("FUNIL DE CONVERSÃO DE CREATORS");
  dashSheet.getRange("A11").setFontWeight("bold").setHorizontalAlignment("center");
  dashSheet.getRange("A11:E11").setBackground("#e3f2fd");
  
  // Área para gráfico de funil
  dashSheet.getRange("A12:E22").merge();
  dashSheet.getRange("A12:E22").setBackground("#ffffff").setBorder(true, true, true, true, true, true);
  
  // Dados para o gráfico de funil
  dashSheet.getRange("G12").setValue("Status");
  dashSheet.getRange("H12").setValue("Quantidade");
  
  var statusArray = [
    ["Não Contatado", "=COUNTIF('Guia Principal'!L2:L, \"Não Contatado\")"],
    ["Contatado", "=COUNTIF('Guia Principal'!L2:L, \"Contatado\")"],
    ["Aguardando Resposta", "=COUNTIF('Guia Principal'!L2:L, \"Aguardando Resposta\")"],
    ["Reunião Agendada", "=COUNTIF('Guia Principal'!L2:L, \"Reunião Agendada\")"],
    ["Confirmado", "=COUNTIF('Guia Principal'!L2:L, \"Confirmado\")"],
    ["Parceria Fechada", "=COUNTIF('Guia Principal'!L2:L, \"Parceria Fechada\")"],
    ["Parceria Ativa", "=COUNTIF('Guia Principal'!L2:L, \"Parceria Ativa\")"]
  ];
  
  for (var i = 0; i < statusArray.length; i++) {
    dashSheet.getRange(13 + i, 7).setValue(statusArray[i][0]);
    dashSheet.getRange(13 + i, 8).setFormula(statusArray[i][1]);
  }
  
  // Instrução para criação do gráfico
  dashSheet.getRange("A23").setValue("Para criar o gráfico de funil: Inserir > Gráfico > Selecione G12:H19 > Tipo: Funil");
  dashSheet.getRange("A23").setFontStyle("italic").setFontSize(10);
  
  // ==== SEÇÃO 2: DISTRIBUIÇÃO POR PLATAFORMA E TIER ====
  dashSheet.getRange("A25:E25").merge();
  dashSheet.getRange("A25").setValue("DISTRIBUIÇÃO POR PLATAFORMA E TIER");
  dashSheet.getRange("A25").setFontWeight("bold").setHorizontalAlignment("center");
  dashSheet.getRange("A25:E25").setBackground("#e3f2fd");
  
  // Área para gráficos de distribuição
  dashSheet.getRange("A26:B36").merge();
  dashSheet.getRange("A26:B36").setBackground("#ffffff").setBorder(true, true, true, true, true, true);
  dashSheet.getRange("A26").setValue("Gráfico: Creators por Plataforma");
  
  dashSheet.getRange("C26:E36").merge();
  dashSheet.getRange("C26:E36").setBackground("#ffffff").setBorder(true, true, true, true, true, true);
  dashSheet.getRange("C26").setValue("Gráfico: Creators por Tier");
  
  // Dados para gráfico de plataformas
  dashSheet.getRange("G25").setValue("Plataforma");
  dashSheet.getRange("H25").setValue("Quantidade");
  
  // Popula os dados para cada plataforma
  for (var i = 0; i < PLATAFORMAS.length; i++) {
    dashSheet.getRange(26 + i, 7).setValue(PLATAFORMAS[i]);
    dashSheet.getRange(26 + i, 8).setFormula("=COUNTIF('Guia Principal'!E2:E, \"" + PLATAFORMAS[i] + "\")");
  }
  
  // Dados para gráfico de Tier
  dashSheet.getRange("J25").setValue("Tier");
  dashSheet.getRange("K25").setValue("Quantidade");
  
  // Popula os dados para cada Tier
  for (var i = 0; i < PRIORIDADE_OPTIONS.length; i++) {
    dashSheet.getRange(26 + i, 10).setValue(PRIORIDADE_OPTIONS[i]);
    dashSheet.getRange(26 + i, 11).setFormula("=COUNTIF('Guia Principal'!Q2:Q, \"" + PRIORIDADE_OPTIONS[i] + "\")");
  }
  
  // Instrução para criação dos gráficos
  dashSheet.getRange("A37").setValue("Para criar os gráficos de distribuição: Inserir > Gráfico > Selecione os dados > Tipo: Pizza ou Barras");
  dashSheet.getRange("A37").setFontStyle("italic").setFontSize(10);
  
  // ==== SEÇÃO 3: ANÁLISE DE DESEMPENHO ====
  dashSheet.getRange("A39:E39").merge();
  dashSheet.getRange("A39").setValue("ANÁLISE DE DESEMPENHO DE CAMPANHAS");
  dashSheet.getRange("A39").setFontWeight("bold").setHorizontalAlignment("center");
  dashSheet.getRange("A39:E39").setBackground("#e3f2fd");
  
  // Área para gráfico de análise
  dashSheet.getRange("A40:E50").merge();
  dashSheet.getRange("A40:E50").setBackground("#ffffff").setBorder(true, true, true, true, true, true);
  
  // Tabela de métricas de desempenho
  dashSheet.getRange("A52:E52").merge();
  dashSheet.getRange("A52").setValue("MÉTRICAS DE ACOMPANHAMENTO");
  dashSheet.getRange("A52").setFontWeight("bold").setHorizontalAlignment("center");
  dashSheet.getRange("A52:E52").setBackground("#e3f2fd");
  
  var metricas = [
    ["Métrica", "Meta", "Realizado", "% Atingido", "Status"],
    ["Creators Captados", 50, "=COUNTA('Guia Principal'!B2:B)", "=C53/B53", "=IF(C53/B53>=1,\"Atingido\",\"Em Progresso\")"],
    ["Reuniões Realizadas", 30, "=COUNTIF('Calendário de Reuniões'!F2:F,\"Realizada\")", "=C54/B54", "=IF(C54/B54>=1,\"Atingido\",\"Em Progresso\")"],
    ["Parcerias Fechadas", 15, "=COUNTIF('Guia Principal'!L2:L,\"Parceria Fechada\")+COUNTIF('Guia Principal'!L2:L,\"Parceria Ativa\")", "=C55/B55", "=IF(C55/B55>=1,\"Atingido\",\"Em Progresso\")"],
    ["Média Engajamento", 0.05, "=AVERAGE('Guia Principal'!G2:G)", "=C56/B56", "=IF(C56/B56>=1,\"Atingido\",\"Em Progresso\")"],
    ["ROI Campanha", 2, "=IFERROR(AVERAGE('Histórico de Campanhas'!G2:G), 0)", "=C57/B57", "=IF(C57/B57>=1,\"Atingido\",\"Em Progresso\")"]
  ];
  
  dashSheet.getRange(31, 2, metricas.length, metricas[0].length).setValues(metricas);
  dashSheet.getRange(31, 2, 1, metricas[0].length).setFontWeight("bold").setBackground("#f5f5f5");
  
  // Formatação para métricas
  dashSheet.getRange("B32:C32").setNumberFormat("R$ #,##0.00");
  dashSheet.getRange("D32").setNumberFormat("0.00%");
  
  // Adiciona botão para atualizar
  dashSheet.getRange("E2").setValue("Atualizar Sumário");
  dashSheet.getRange("E2").setFontWeight("bold").setBackground("#e3f2fd");
}

/**
 * Aplica regras de validação nas guias
 */
function aplicarRegraValidacao() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  
  // ======== Guia Principal ========
  var principalSheet = ss.getSheetByName("Guia Principal");
  
  // Regra para Status Contato (coluna 12)
  var statusRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(STATUS_OPTIONS, true)
    .setAllowInvalid(false)
    .build();
  principalSheet.getRange("L2:L1000").setDataValidation(statusRule);
  
  // Regra para Prioridade (coluna 17)
  var prioridadeRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(PRIORIDADE_OPTIONS, true)
    .setAllowInvalid(false)
    .build();
  principalSheet.getRange("Q2:Q1000").setDataValidation(prioridadeRule);
  
  // Regra para Etapa Jornada (coluna 18)
  var etapaRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(ETAPAS_JORNADA, true)
    .setAllowInvalid(false)
    .build();
  principalSheet.getRange("R2:R1000").setDataValidation(etapaRule);
  
  // Regra para Nível Personalização (coluna 19)
  var nivelRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(NIVEL_PERSONALIZACAO, true)
    .setAllowInvalid(false)
    .build();
  principalSheet.getRange("S2:S1000").setDataValidation(nivelRule);
  
  // Regra para Plataforma Principal (coluna 5)
  var plataformaRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(PLATAFORMAS, true)
    .setAllowInvalid(false)
    .build();
  principalSheet.getRange("E2:E1000").setDataValidation(plataformaRule);
  
  // Formatação de data para Data Primeiro Contato (coluna 11)
  principalSheet.getRange("K2:K1000").setNumberFormat("dd/mm/yyyy");
  
  // Formatação de data para Data Reunião (coluna 14)
  principalSheet.getRange("N2:N1000").setNumberFormat("dd/mm/yyyy");
  
  // Formatação de percentual para Engajamento (coluna 7)
  principalSheet.getRange("G2:G1000").setNumberFormat("0.00%");
  
  // Formatação para números de seguidores (coluna 6)
  principalSheet.getRange("F2:F1000").setNumberFormat("#,##0");
  
  // ======== Guia Qualificação dos Creators ========
  var qualSheet = ss.getSheetByName("Qualificação dos Creators");
  
  // Validação para pontuações (1-5)
  var rangeValidation = SpreadsheetApp.newDataValidation()
    .requireNumberBetween(1, 5)
    .setAllowInvalid(false)
    .build();
  qualSheet.getRange("B2:G1000").setDataValidation(rangeValidation);
  
  // Validação para Segmento
  var segmentoRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(SEGMENTO_OPTIONS, true)
    .setAllowInvalid(false)
    .build();
  qualSheet.getRange("J2:J1000").setDataValidation(segmentoRule);
  
  // Validação para Tier de Prioridade
  var tierRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(PRIORIDADE_OPTIONS, true)
    .setAllowInvalid(false)
    .build();
  qualSheet.getRange("K2:K1000").setDataValidation(tierRule);
  
  // ======== Guia ICP de Creators ========
  var icpSheet = ss.getSheetByName("ICP de Creators");
  if (icpSheet) {
    // Preenche critérios padrão para ICP se a planilha estiver vazia
    var icpData = icpSheet.getDataRange().getValues();
    if (icpData.length <= 1) { // Só tem cabeçalho ou vazio
      var criterios = [
        ["Critério", "Perfil Ideal (ICP)", "Creator 1", "Creator 2", "Creator 3", "Creator 4", "Creator 5", "Creator 6", "Creator 7", "Creator 8"],
        ["Tamanho da Audiência", "", "", "", "", "", "", "", "", ""],
        ["Engajamento Médio", "", "", "", "", "", "", "", "", ""],
        ["Nicho de Mercado", "", "", "", "", "", "", "", "", ""],
        ["Alinhamento com Marca", "", "", "", "", "", "", "", "", ""],
        ["Qualidade do Conteúdo", "", "", "", "", "", "", "", "", ""],
        ["Exclusividade", "", "", "", "", "", "", "", "", ""],
        ["Credibilidade", "", "", "", "", "", "", "", "", ""],
        ["ROI Potencial", "", "", "", "", "", "", "", "", ""],
        ["Taxa de Crescimento", "", "", "", "", "", "", "", "", ""],
        ["Disponibilidade", "", "", "", "", "", "", "", "", ""]
      ];
      icpSheet.getRange(1, 1, criterios.length, criterios[0].length).setValues(criterios);
      
      // Formata o cabeçalho
      icpSheet.getRange(1, 1, 1, criterios[0].length).setBackground("#4285F4").setFontColor("white").setFontWeight("bold");
    }
  }
  
  // ======== Guia Plano de Conteúdo ========
  var contentSheet = ss.getSheetByName("Plano de Conteúdo");
  if (contentSheet) {
    // Validação para Formato de Conteúdo
    var formatoRule = SpreadsheetApp.newDataValidation()
      .requireValueInList(FORMATO_CONTEUDO, true)
      .setAllowInvalid(false)
      .build();
    contentSheet.getRange("H2:H1000").setDataValidation(formatoRule);
    
    // Validação para Status de Conteúdo
    var statusContentRule = SpreadsheetApp.newDataValidation()
      .requireValueInList(STATUS_CONTEUDO, true)
      .setAllowInvalid(false)
      .build();
    contentSheet.getRange("J2:J1000").setDataValidation(statusContentRule);
    
    // Formatação de datas
    contentSheet.getRange("I2:I1000").setNumberFormat("dd/mm/yyyy");
  }
  
  // ======== Guia Calendário Editorial ========
  var calendarSheet = ss.getSheetByName("Calendário Editorial");
  if (calendarSheet) {
    // Validação para Formato
    calendarSheet.getRange("D2:D1000").setDataValidation(formatoRule);
    
    // Validação para Plataforma
    calendarSheet.getRange("E2:E1000").setDataValidation(plataformaRule);
    
    // Validação para Status
    calendarSheet.getRange("G2:G1000").setDataValidation(statusContentRule);
    
    // Formatação de datas
    calendarSheet.getRange("F2:F1000").setNumberFormat("dd/mm/yyyy");
  }
  
  // ======== Guia Métricas de Desempenho ========
  var metricsSheet = ss.getSheetByName("Métricas de Desempenho");
  if (metricsSheet) {
    // Preencher com métricas padrão se estiver vazia
    var metricsData = metricsSheet.getDataRange().getValues();
    if (metricsData.length <= 1) {
      var metricas = [
        ["Categoria", "Tipo", "Métrica", "Checkpoint 1", "Checkpoint 2", "Checkpoint 3", "Checkpoint 4", "Checkpoint 5", "Meta", "Realizado", "Status"],
        ["Engajamento", "Cobertura", "Número de contatos de alto nível", "", "", "", "", "", "", "", ""],
        ["Engajamento", "Cobertura", "Número de departamentos influenciados", "", "", "", "", "", "", "", ""],
        ["Engajamento", "Atividade", "Número de interações", "", "", "", "", "", "", "", ""],
        ["Engajamento", "Atividade", "Cliques/Visualizações", "", "", "", "", "", "", "", ""],
        ["Engajamento", "Engajamento", "Taxa de engajamento", "", "", "", "", "", "", "", ""],
        ["Pipeline", "Conversão", "Número de oportunidades geradas", "", "", "", "", "", "", "", ""],
        ["Pipeline", "Conversão", "Taxa de conversão", "", "", "", "", "", "", "", ""],
        ["Pipeline", "Valor", "Valor total de oportunidades", "", "", "", "", "", "", "", ""],
        ["Pipeline", "Velocidade", "Tempo médio de ciclo", "", "", "", "", "", "", "", ""],
        ["Resultados", "Receita", "Receita total gerada", "", "", "", "", "", "", "", ""],
        ["Resultados", "ROI", "Retorno sobre investimento", "", "", "", "", "", "", "", ""]
      ];
      metricsSheet.getRange(1, 1, metricas.length, metricas[0].length).setValues(metricas);
      
      // Formata o cabeçalho
      metricsSheet.getRange(1, 1, 1, metricas[0].length).setBackground("#4285F4").setFontColor("white").setFontWeight("bold");
      
      // Validação para colunas de Checkpoint
      var checkpointRule = SpreadsheetApp.newDataValidation()
        .requireValueInList(CHECKPOINT_OPTIONS, true)
        .setAllowInvalid(false)
        .build();
      metricsSheet.getRange("D1:H1").setDataValidation(checkpointRule);
    }
  }
  
  // ======== Guia Alocação de Recursos ========
  var resourceSheet = ss.getSheetByName("Alocação de Recursos");
  if (resourceSheet) {
    // Preencher com categorias padrão se estiver vazia
    var resourceData = resourceSheet.getDataRange().getValues();
    if (resourceData.length <= 1) {
      var recursos = [
        ["Categoria", "Orçamento Estimado", "Orçamento Realizado", "Variação", "Observações"],
        ["Produção de Conteúdo", "", "", "=B2-C2", ""],
        ["Cachê de Creators", "", "", "=B3-C3", ""],
        ["Impulsionamento", "", "", "=B4-C4", ""],
        ["Eventos com Creators", "", "", "=B5-C5", ""],
        ["Ferramentas de Gestão", "", "", "=B6-C6", ""],
        ["Comissões", "", "", "=B7-C7", ""],
        ["Logística e Envios", "", "", "=B8-C8", ""],
        ["Outros", "", "", "=B9-C9", ""],
        ["Total", "=SUM(B2:B9)", "=SUM(C2:C9)", "=B10-C10", ""]
      ];
      resourceSheet.getRange(1, 1, recursos.length, recursos[0].length).setValues(recursos);
      
      // Formata o cabeçalho
      resourceSheet.getRange(1, 1, 1, recursos[0].length).setBackground("#4285F4").setFontColor("white").setFontWeight("bold");
      
      // Formata a linha de total
      resourceSheet.getRange(10, 1, 1, recursos[0].length).setBackground("#e3f2fd").setFontWeight("bold");
      
      // Formata colunas de valor como moeda
      resourceSheet.getRange("B2:D10").setNumberFormat("R$ #,##0.00");
    }
  }
  
  // ======== Guia Alinhamento Estratégico ========
  var strategySheet = ss.getSheetByName("Alinhamento Estratégico");
  if (strategySheet) {
    // Validação para Status
    var statusProjRule = SpreadsheetApp.newDataValidation()
      .requireValueInList(["Não Iniciado", "Em Andamento", "Atrasado", "Concluído", "Cancelado"], true)
      .setAllowInvalid(false)
      .build();
    strategySheet.getRange("G2:G1000").setDataValidation(statusProjRule);
  }
}

/**
 * FUNÇÕES DE MANIPULAÇÃO DE DADOS
 */

/**
 * Manipula eventos de edição na planilha com tratamento de erros
 */
function onEdit(e) {
  try {
    var sheet = e.source.getActiveSheet();
    var range = e.range;
    var row = range.getRow();
    var col = range.getColumn();
    
    // Gerencia eventos na Guia Principal
    if(sheet.getName() == "Guia Principal"){
      // Quando o Status é alterado
      if(col == 12){ // Coluna 12 = Status Contato
        var status = range.getValue();
        var actionCell = sheet.getRange(row, 13); // Coluna 13 = Próxima Ação
        
        if(status == "Confirmado"){
          actionCell.setValue("Marcar Reunião");
        } else if(status == "Aguardando Resposta"){
          actionCell.setValue("Enviar Follow-up");
        } else if(status == "Reunião Agendada"){
          actionCell.setValue("Preparar Material");
          // Destaca visualmente
          sheet.getRange(row, 1, 1, sheet.getLastColumn()).setBackground("#fffde7");
        } else if(status == "Parceria Fechada"){
          actionCell.setValue("Iniciar Onboarding");
          // Destaca visualmente
          sheet.getRange(row, 1, 1, sheet.getLastColumn()).setBackground("#e8f5e9");
        } else if(status == "Parceria Ativa"){
          actionCell.setValue("Monitorar Resultados");
          // Destaca visualmente
          sheet.getRange(row, 1, 1, sheet.getLastColumn()).setBackground("#e8f5e9");
        } else if(status == "Não Interessado"){
          actionCell.setValue("Arquivar");
          // Diminui visualmente
          sheet.getRange(row, 1, 1, sheet.getLastColumn()).setBackground("#f5f5f5");
        } else {
          actionCell.clearContent();
        }
      }
      
      // Quando uma data de reunião é inserida
      if(col == 14){ // Coluna 14 = Data Reunião
        var dataReuniao = range.getValue();
        if(dataReuniao){
          // Atualiza a guia Calendário de Reuniões
          var calendarSheet = e.source.getSheetByName("Calendário de Reuniões");
          if (calendarSheet) {
            var nome = sheet.getRange(row, 2).getValue(); // Nome do Creator
            var horario = sheet.getRange(row, 15).getValue() || ""; // Horário Reunião
            
            // Adiciona a reunião ao calendário
            calendarSheet.appendRow([dataReuniao, horario, nome, "", "", "Agendada"]);
            
            // Atualiza o status se necessário
            var statusCell = sheet.getRange(row, 12);
            if(statusCell.getValue() != "Parceria Fechada" && statusCell.getValue() != "Parceria Ativa"){
              statusCell.setValue("Reunião Agendada");
            }
          }
        }
      }
      
      // Gera IDs automaticamente quando um novo nome é adicionado
      if(col == 2 && row > 1){ // Coluna 2 = Nome
        var idCell = sheet.getRange(row, 1);
        if(!idCell.getValue() && range.getValue()){
          idCell.setValue("CR" + new Date().getTime().toString().slice(-6));
        }
      }
      
      // Calcula pontuação na guia de Qualificação quando adicionado um novo creator
      if(col == 2 && row > 1 && range.getValue()){ // Coluna 2 = Nome
        var nome = range.getValue();
        var qualSheet = e.source.getSheetByName("Qualificação dos Creators");
        
        if (qualSheet) {
          // Verifica se o creator já existe na guia de qualificação
          var qualData = qualSheet.getDataRange().getValues();
          var creatorExiste = false;
          
          for(var i = 1; i < qualData.length; i++){
            if(qualData[i][0] == nome){
              creatorExiste = true;
              break;
            }
          }
          
          // Se não existe, adiciona à guia de qualificação
          if(!creatorExiste){
            qualSheet.appendRow([nome, "", "", "", "", "", "", "=SUM(B" + (qualSheet.getLastRow() + 1) + ":G" + (qualSheet.getLastRow() + 1) + ")", "", "", ""]);
          }
        }
      }
    }
    
    // Gerencia eventos na guia Qualificação dos Creators
    if(sheet.getName() == "Qualificação dos Creators"){
      // Atualiza pontuação total quando qualquer pontuação é alterada
      if(col >= 2 && col <= 7 && row > 1){ // Colunas de pontuação (B-G)
        var pontuacaoCell = sheet.getRange(row, 8); // Coluna H = Pontuação Total
        var statusCell = sheet.getRange(row, 9); // Coluna I = Status Qualificação
        var segmentoCell = sheet.getRange(row, 10); // Coluna J = Segmento
        var tierCell = sheet.getRange(row, 11); // Coluna K = Tier de Prioridade
        
        // Verifica se a fórmula de soma está presente
        if(!pontuacaoCell.getFormula()){
          pontuacaoCell.setFormula("=SUM(B" + row + ":G" + row + ")");
        }
        
        // Define o status com base na pontuação total
        var pontuacao = pontuacaoCell.getValue();
        if(pontuacao >= 20){
          statusCell.setValue("Excelente (20-25)");
          sheet.getRange(row, 8, 1, 2).setBackground("#c8e6c9"); // Verde claro
          
          // Define tier com base na pontuação
          if (!tierCell.getValue()) {
            tierCell.setValue("Nível 1 (Alta)");
          }
          
          // Define segmento se estiver vazio
          if (!segmentoCell.getValue()) {
            segmentoCell.setValue("Estratégico");
          }
        } else if(pontuacao >= 15){
          statusCell.setValue("Bom (15-19)");
          sheet.getRange(row, 8, 1, 2).setBackground("#fff9c4"); // Amarelo claro
          
          // Define tier com base na pontuação
          if (!tierCell.getValue()) {
            tierCell.setValue("Nível 2 (Média)");
          }
          
          // Define segmento se estiver vazio
          if (!segmentoCell.getValue()) {
            segmentoCell.setValue("Crescimento");
          }
        } else if(pontuacao >= 10){
          statusCell.setValue("Regular (10-14)");
          sheet.getRange(row, 8, 1, 2).setBackground("#ffecb3"); // Laranja claro
          
          // Define tier com base na pontuação
          if (!tierCell.getValue()) {
            tierCell.setValue("Nível 3 (Baixa)");
          }
          
          // Define segmento se estiver vazio
          if (!segmentoCell.getValue()) {
            segmentoCell.setValue("Oportunidade");
          }
        } else {
          statusCell.setValue("Baixo Potencial (<10)");
          sheet.getRange(row, 8, 1, 2).setBackground("#ffcdd2"); // Vermelho claro
          
          // Define tier com base na pontuação
          if (!tierCell.getValue()) {
            tierCell.setValue("Lista de Espera");
          }
        }
      }
    }
    
    // Gerencia eventos na guia Calendário de Reuniões
    if(sheet.getName() == "Calendário de Reuniões"){
      // Quando o status da reunião é alterado
      if(col == 6){ // Coluna 6 = Status da Reunião
        var status = range.getValue();
        if(status == "Realizada"){
          // Destaca visualmente
          sheet.getRange(row, 1, 1, sheet.getLastColumn()).setBackground("#e8f5e9");
        } else if(status == "Cancelada"){
          // Destaca visualmente
          sheet.getRange(row, 1, 1, sheet.getLastColumn()).setBackground("#ffcdd2");
        } else if(status == "Reagendada"){
          // Destaca visualmente
          sheet.getRange(row, 1, 1, sheet.getLastColumn()).setBackground("#fff9c4");
        }
      }
    }
    
    // Gerencia eventos na guia Histórico de Campanhas
    if(sheet.getName() == "Histórico de Campanhas"){
      // Calcula ROI quando resultado é adicionado
      if(col == 5 && row > 1){ // Coluna 5 = Resultado
        var resultado = range.getValue();
        if(resultado && !isNaN(parseFloat(resultado))){
          // Adiciona fórmula para calcular ROI
          sheet.getRange(row, 7).setFormula("=E" + row + "/1000"); // Exemplo de cálculo simplificado
        }
      }
    }
    
    // Atualiza estatísticas quando há mudanças relevantes nas guias principais
    if(["Guia Principal", "Qualificação dos Creators", "Histórico de Campanhas"].includes(sheet.getName())){
      // Atualiza as estatísticas, mas limita a frequência para não sobrecarregar
      var cache = CacheService.getScriptCache();
      var lastUpdate = cache.get('lastStatsUpdate');
      var now = new Date().getTime();
      
      if(!lastUpdate || (now - parseInt(lastUpdate)) > 10000){ // 10 segundos entre atualizações
        atualizarEstatisticas();
        cache.put('lastStatsUpdate', now.toString(), 60); // Cache por 1 minuto
      }
    }
  } catch(e) {
    // Log silencioso de erros para não interferir na experiência do usuário
    console.error("Erro no onEdit: " + e.toString());
  }
}

/**
 * Atualiza estatísticas e métricas em todas as planilhas
 */
function atualizarEstatisticas() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  
  // Força atualização das estatísticas gerais
  var statsSheet = ss.getSheetByName("Estatísticas Gerais");
  if (statsSheet) {
    // Recalcula as fórmulas forçando refresh
    var totalCreators = statsSheet.getRange("B2").getValue();
    statsSheet.getRange("B2").setValue(totalCreators);
  }
  
  // Atualiza o painel de controle
  var dashSheet = ss.getSheetByName("Painel de Controle");
  if (dashSheet) {
    // Força recálculo das fórmulas
    var metricas = dashSheet.getRange("A7:C7").getValues();
    dashSheet.getRange("A7:C7").setValues(metricas);
  }
  
  // Atualiza o sumário executivo
  var sumarioSheet = ss.getSheetByName("Sumário Executivo");
  if (sumarioSheet) {
    var kpis = sumarioSheet.getRange("B12:D16").getValues();
    sumarioSheet.getRange("B12:D16").setValues(kpis);
  }
  
  // Exibe mensagem
  SpreadsheetApp.getUi().alert("Estatísticas atualizadas com sucesso!");
}

/**
 * FUNÇÕES DE EXPORTAÇÃO E RELATÓRIOS
 */

/**
 * Exporta o calendário de reuniões para o Google Calendar
 * Requer autorização adicional para o Google Calendar
 */
function exportarCalendario() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var calSheet = ss.getSheetByName("Calendário de Reuniões");
  
  if(!calSheet) {
    SpreadsheetApp.getUi().alert("Guia de Calendário não encontrada!");
    return;
  }
  
  var data = calSheet.getDataRange().getValues();
  var ui = SpreadsheetApp.getUi();
  
  try {
    var calendar = CalendarApp.getDefaultCalendar();
    var contador = 0;
    
    // Começa da linha 1 (pula o cabeçalho)
    for(var i = 1; i < data.length; i++) {
      var dataReuniao = data[i][0]; // Data
      var horario = data[i][1]; // Horário
      var creator = data[i][2]; // Nome Creator
      var plataforma = data[i][3]; // Plataforma
      var link = data[i][4]; // Link
      
      // Verifica se tem data e hora válidas
      if(dataReuniao && horario) {
        // Cria uma data completa concatenando data e hora
        var dataCompleta = new Date(dataReuniao);
        var horaParts = horario.split(":");
        if(horaParts.length >= 2) {
          dataCompleta.setHours(parseInt(horaParts[0]), parseInt(horaParts[1]));
          
          // Cria evento de 1 hora
          var endTime = new Date(dataCompleta.getTime() + 60 * 60 * 1000);
          
          // Descrição do evento
          var desc = "Reunião com creator: " + creator;
          if(plataforma) desc += "\nPlataforma: " + plataforma;
          if(link) desc += "\nLink: " + link;
          
          // Cria o evento
          calendar.createEvent(
            "Reunião: " + creator,
            dataCompleta,
            endTime,
            {description: desc}
          );
          
          contador++;
        }
      }
    }
    
    ui.alert("Calendário exportado com sucesso! " + contador + " eventos criados.");
  } catch(e) {
    ui.alert("Erro ao exportar para o calendário: " + e.toString());
  }
}

/**
 * Gerar relatório de desempenho
 */
function gerarRelatorio() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var mainSheet = ss.getSheetByName("Guia Principal");
  var reportSheet = ss.getSheetByName("Relatório");
  
  if (!reportSheet) {
    reportSheet = ss.insertSheet("Relatório");
  }
  
  reportSheet.clear();
  
  // Configuração básica da planilha
  reportSheet.setColumnWidth(1, 30);
  reportSheet.setColumnWidth(2, 200);
  reportSheet.setColumnWidth(3, 200);
  reportSheet.setColumnWidth(4, 200);
  reportSheet.setColumnWidth(5, 200);
  reportSheet.setColumnWidth(6, 200);
  reportSheet.setColumnWidth(7, 200);
  
  // Título do relatório
  reportSheet.getRange("A1:G1").merge();
  reportSheet.getRange("A1").setValue("RELATÓRIO DE DESEMPENHO - PROGRAMA DE CREATORS");
  reportSheet.getRange("A1").setFontSize(16).setFontWeight("bold").setHorizontalAlignment("center");
  reportSheet.getRange("A1:G1").setBackground("#4285F4").setFontColor("white");
  
  // Data do relatório
  var today = new Date();
  reportSheet.getRange("A2:G2").merge();
  reportSheet.getRange("A2").setValue("Período: " + Utilities.formatDate(today, "GMT-3", "MMMM yyyy"));
  reportSheet.getRange("A2").setHorizontalAlignment("center").setFontStyle("italic");
  
  // ===== SEÇÃO 1: VISÃO GERAL =====
  var nextRow = 4;
  reportSheet.getRange(nextRow, 1, 1, 7).merge();
  reportSheet.getRange(nextRow, 1).setValue("1. VISÃO GERAL");
  reportSheet.getRange(nextRow, 1).setFontWeight("bold").setFontSize(12);
  reportSheet.getRange(nextRow, 1, 1, 7).setBackground("#e3f2fd");
  
  // Métricas principais
  var metricas = [
    ["Total de Creators", "=COUNTA('Guia Principal'!B2:B)"],
    ["Parcerias Ativas", "=COUNTIF('Guia Principal'!L2:L, \"Parceria Ativa\")"],
    ["Taxa de Conversão", "=COUNTIF('Guia Principal'!L2:L, \"Parceria Fechada\")/COUNTA('Guia Principal'!B2:B)"],
    ["Média de Engajamento", "=AVERAGE('Guia Principal'!G2:G)"],
    ["ROI Médio", "=IFERROR(AVERAGE('Histórico de Campanhas'!G2:G), 0)"]
  ];
  
  for (var i = 0; i < metricas.length; i++) {
    reportSheet.getRange(nextRow + 1 + i, 1).setValue(metricas[i][0]);
    reportSheet.getRange(nextRow + 1 + i, 2).setFormula(metricas[i][1]);
  }
  
  // ===== SEÇÃO 2: DISTRIBUIÇÃO POR PLATAFORMA =====
  nextRow = nextRow + metricas.length + 3;
  reportSheet.getRange(nextRow, 1, 1, 7).merge();
  reportSheet.getRange(nextRow, 1).setValue("2. DISTRIBUIÇÃO POR PLATAFORMA");
  reportSheet.getRange(nextRow, 1).setFontWeight("bold").setFontSize(12);
  reportSheet.getRange(nextRow, 1, 1, 7).setBackground("#e3f2fd");
  
  // Tabela de plataformas
  reportSheet.getRange(nextRow + 1, 1).setValue("Plataforma");
  reportSheet.getRange(nextRow + 1, 2).setValue("Quantidade");
  reportSheet.getRange(nextRow + 1, 3).setValue("% do Total");
  
  var plataformas = PLATAFORMAS;
  for (var i = 0; i < plataformas.length; i++) {
    reportSheet.getRange(nextRow + 2 + i, 1).setValue(plataformas[i]);
    reportSheet.getRange(nextRow + 2 + i, 2).setFormula("=COUNTIF('Guia Principal'!E2:E, \"" + plataformas[i] + "\")");
    reportSheet.getRange(nextRow + 2 + i, 3).setFormula("=IF(B" + (nextRow + 2 + i) + ">0, B" + (nextRow + 2 + i) + "/B" + (nextRow + 1) + ", 0)");
  }
  
  // Formatação para percentuais
  reportSheet.getRange(nextRow + 2, 3, plataformas.length).setNumberFormat("0.00%");
  
  // ===== SEÇÃO 3: DESEMPENHO POR TIER =====
  nextRow = nextRow + plataformas.length + 3;
  reportSheet.getRange(nextRow, 1, 1, 7).merge();
  reportSheet.getRange(nextRow, 1).setValue("3. DESEMPENHO POR TIER");
  reportSheet.getRange(nextRow, 1).setFontWeight("bold").setFontSize(12);
  reportSheet.getRange(nextRow, 1, 1, 7).setBackground("#e3f2fd");
  
  // Tabela de tiers
  var tiers = ["Nível 1 (Alta)", "Nível 2 (Média)", "Nível 3 (Baixa)", "Lista de Espera"];
  reportSheet.getRange(nextRow + 1, 1).setValue("Tier");
  reportSheet.getRange(nextRow + 1, 2).setValue("Quantidade");
  reportSheet.getRange(nextRow + 1, 3).setValue("Taxa de Conversão");
  reportSheet.getRange(nextRow + 1, 4).setValue("Média Engajamento");
  
  for (var i = 0; i < tiers.length; i++) {
    reportSheet.getRange(nextRow + 2 + i, 1).setValue(tiers[i]);
    reportSheet.getRange(nextRow + 2 + i, 2).setFormula("=COUNTIF('Guia Principal'!Q2:Q, \"" + tiers[i] + "\")");
    reportSheet.getRange(nextRow + 2 + i, 3).setFormula("=COUNTIFS('Guia Principal'!Q2:Q, \"" + tiers[i] + "\", 'Guia Principal'!L2:L, \"Parceria Fechada\")/B" + (nextRow + 2 + i));
    reportSheet.getRange(nextRow + 2 + i, 4).setFormula("=AVERAGEIF('Guia Principal'!Q2:Q, \"" + tiers[i] + "\", 'Guia Principal'!G2:G)");
  }
  
  // Formatação para percentuais
  reportSheet.getRange(nextRow + 2, 3, tiers.length).setNumberFormat("0.00%");
  reportSheet.getRange(nextRow + 2, 4, tiers.length).setNumberFormat("0.00%");
  
  // ===== SEÇÃO 4: CAMPANHAS RECENTES =====
  nextRow = nextRow + tiers.length + 3;
  reportSheet.getRange(nextRow, 1, 1, 7).merge();
  reportSheet.getRange(nextRow, 1).setValue("4. CAMPANHAS RECENTES");
  reportSheet.getRange(nextRow, 1).setFontWeight("bold").setFontSize(12);
  reportSheet.getRange(nextRow, 1, 1, 7).setBackground("#e3f2fd");
  
  // Tabela de campanhas
  reportSheet.getRange(nextRow + 1, 1).setValue("Data");
  reportSheet.getRange(nextRow + 1, 2).setValue("Creator");
  reportSheet.getRange(nextRow + 1, 3).setValue("Tipo");
  reportSheet.getRange(nextRow + 1, 4).setValue("Resultado");
  reportSheet.getRange(nextRow + 1, 5).setValue("ROI");
  
  // Busca as últimas 5 campanhas
  var campSheet = ss.getSheetByName("Histórico de Campanhas");
  if (campSheet) {
    var campData = campSheet.getDataRange().getValues();
    var campRow = Math.min(6, campData.length);
    
    for (var i = 1; i < campRow; i++) {
      reportSheet.getRange(nextRow + 1 + i, 1).setValue(campData[i][2]); // Data
      reportSheet.getRange(nextRow + 1 + i, 2).setValue(campData[i][1]); // Creator
      reportSheet.getRange(nextRow + 1 + i, 3).setValue(campData[i][3]); // Tipo
      reportSheet.getRange(nextRow + 1 + i, 4).setValue(campData[i][4]); // Resultado
      reportSheet.getRange(nextRow + 1 + i, 5).setValue(campData[i][6]); // ROI
    }
  }
  
  // Formatação para ROI
  reportSheet.getRange(nextRow + 2, 5, 4).setNumberFormat("0.00%");
  
  // ===== SEÇÃO 5: PRINCIPAIS INSIGHTS E RECOMENDAÇÕES =====
  nextRow = Math.max(campRow + 2, nextRow + 12);
  
  reportSheet.getRange(nextRow, 1, 1, 7).merge();
  reportSheet.getRange(nextRow, 1).setValue("5. INSIGHTS E RECOMENDAÇÕES");
  reportSheet.getRange(nextRow, 1).setFontWeight("bold").setFontSize(12);
  reportSheet.getRange(nextRow, 1, 1, 7).setBackground("#e3f2fd");
  
  reportSheet.getRange(nextRow + 1, 1, 1, 7).merge();
  reportSheet.getRange(nextRow + 1, 1).setValue("Os insights e recomendações abaixo são baseados na análise dos dados do período:");
  reportSheet.getRange(nextRow + 1, 1).setFontStyle("italic");
  
  // Adiciona linhas para insights
  var insightsTexto = [
    "• Distribuição por plataforma: " + getTopPlataforma(mainSheet) + " é a plataforma com maior número de creators.",
    "• Qualificação: " + getTopQualificacao(ss) + " dos creators avaliados possuem pontuação 'Excelente'.",
    "• Taxa de conversão: " + getTaxaConversao(mainSheet) + " dos creators contatados avançam para uma parceria.",
    "• Engajamento: A média de engajamento dos creators ativos é de " + getMediaEngajamento(mainSheet) + "%.",
    "",
    "Recomendações:",
    "1. [Incluir recomendações baseadas nos dados, como focar em determinadas plataformas ou tipos de creators]",
    "2. [Recomendação de ajuste na estratégia com base na análise dos resultados]",
    "3. [Sugestões para melhorar a taxa de conversão ou outros KPIs relevantes]"
  ];
  
  for (var i = 0; i < insightsTexto.length; i++) {
    reportSheet.getRange(nextRow + 2 + i, 1, 1, 7).merge();
    reportSheet.getRange(nextRow + 2 + i, 1).setValue(insightsTexto[i]);
  }
  
  // Ajusta largura das colunas
  reportSheet.autoResizeColumns(1, 7);
  
  // Protege o relatório contra edições acidentais
  var protection = reportSheet.protect().setDescription("Proteger Relatório");
  protection.setWarningOnly(true);
  
  // Exibe mensagem
  SpreadsheetApp.getUi().alert("Relatório gerado com sucesso!");
}

/**
 * FUNÇÕES UTILITÁRIAS
 */

/**
 * Funções auxiliares para o relatório
 */
function getWeekNumber(d) {
  d = new Date(Date.UTC(d.getFullYear(), d.getMonth(), d.getDate()));
  d.setUTCDate(d.getUTCDate() + 4 - (d.getUTCDay()||7));
  var yearStart = new Date(Date.UTC(d.getUTCFullYear(),0,1));
  var weekNo = Math.ceil(( ( (d - yearStart) / 86400000) + 1)/7);
  return weekNo;
}

function getTopPlataforma(sheet) {
  var data = sheet.getDataRange().getValues();
  var plataformas = {};
  
  // Coluna 4 é a plataforma principal
  for (var i = 1; i < data.length; i++) {
    var plataforma = data[i][4]; // Coluna E (índice 4)
    if (plataforma) {
      if (!plataformas[plataforma]) {
        plataformas[plataforma] = 0;
      }
      plataformas[plataforma]++;
    }
  }
  
  var topPlataforma = "";
  var maxCount = 0;
  
  for (var plat in plataformas) {
    if (plataformas[plat] > maxCount) {
      maxCount = plataformas[plat];
      topPlataforma = plat;
    }
  }
  
  return topPlataforma;
}

function getTopQualificacao(ss) {
  var qualSheet = ss.getSheetByName("Qualificação dos Creators");
  if (!qualSheet) return "N/A";
  
  var data = qualSheet.getDataRange().getValues();
  var excelentes = 0;
  var total = 0;
  
  // Coluna 7 é a pontuação total, 8 é o status
  for (var i = 1; i < data.length; i++) {
    if (data[i][0]) { // Se tem um nome
      total++;
      if (data[i][8] == "Excelente" || data[i][8] == "Excelente (20-25)") {
        excelentes++;
      }
    }
  }
  
  if (total === 0) return "0%";
  return Math.round(excelentes / total * 100) + "%";
}

function getTaxaConversao(sheet) {
  var data = sheet.getDataRange().getValues();
  var contatados = 0;
  var parcerias = 0;
  
  // Coluna 11 é o status
  for (var i = 1; i < data.length; i++) {
    if (data[i][11] && data[i][11] != "Não Contatado") { // Se foi contatado
      contatados++;
      if (data[i][11] == "Parceria Fechada" || data[i][11] == "Parceria Ativa") {
        parcerias++;
      }
    }
  }
  
  if (contatados === 0) return "0%";
  return Math.round(parcerias / contatados * 100) + "%";
}

function getMediaEngajamento(sheet) {
  var data = sheet.getDataRange().getValues();
  var engajamentos = [];
  
  // Coluna 6 é o engajamento, coluna 11 é o status
  for (var i = 1; i < data.length; i++) {
    if ((data[i][11] == "Parceria Fechada" || data[i][11] == "Parceria Ativa") && data[i][6]) {
      engajamentos.push(data[i][6]);
    }
  }
  
  if (engajamentos.length === 0) return "0";
  
  var soma = 0;
  for (var i = 0; i < engajamentos.length; i++) {
    soma += engajamentos[i];
  }
  
  return (soma / engajamentos.length * 100).toFixed(2);
}

/**
 * FUNÇÕES DE EXEMPLO E DEMONSTRAÇÃO
 */

function criarCreatorExemplo() {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    
    // Adiciona dados na Guia Principal
    var mainSheet = ss.getSheetByName("Guia Principal");
    var mainNextRow = mainSheet.getLastRow() + 1;
    
    var mainData = [
      "CR" + new Date().getTime().toString().slice(-6), // ID
      "TechReview Pro", // Nome
      "Tecnologia", // Categoria
      "Reviews de Produtos", // Nicho Específico
      "YouTube", // Plataforma Principal
      150000, // Seguidores
      0.08, // Engajamento (%)
      "techreview@example.com", // Email
      "+5511999999999", // WhatsApp
      "https://youtube.com/techreviewpro", // Link Perfil
      new Date(), // Data Primeiro Contato
      "Confirmado", // Status Contato
      "Marcar Reunião", // Próxima Ação
      new Date(new Date().setDate(new Date().getDate() + 5)), // Data Reunião
      "14:00", // Horário Reunião
      "João Silva", // Responsável Interno
      "Nível 1 (Alta)", // Prioridade
      "Qualificação", // Etapa Jornada Compra
      "Avançado", // Nível Personalização
      "Creator especializado em reviews de produtos tecnológicos" // Observações
    ];
    
    mainSheet.getRange(mainNextRow, 1, 1, mainData.length).setValues([mainData]);
    
    // Adiciona dados na Qualificação dos Creators
    var qualSheet = ss.getSheetByName("Qualificação dos Creators");
    var qualNextRow = qualSheet.getLastRow() + 1;
    
    var qualData = [
      "TechReview Pro", // Nome Creator
      5, // Potencial Alcance (1-5)
      5, // Relevância para Audiência (1-5)
      5, // Compatibilidade com Marca (1-5)
      5, // Qualidade de Conteúdo (1-5)
      5, // Engajamento da Audiência (1-5)
      5, // Histórico de Conversão (1-5)
      "=SUM(B" + qualNextRow + ":G" + qualNextRow + ")", // Pontuação Total
      "=IF(H" + qualNextRow + ">=20, \"Excelente (20-25)\", IF(H" + qualNextRow + ">=15, \"Bom (15-19)\", IF(H" + qualNextRow + ">=10, \"Regular (10-14)\", \"Baixo Potencial (<10)\")))", // Status Qualificação
      "Estratégico", // Segmento
      "Nível 1 (Alta)" // Nível de Prioridade
    ];
    
    qualSheet.getRange(qualNextRow, 1, 1, qualData.length).setValues([qualData]);
    
    // Adiciona dados no Perfil Detalhado do Creator
    var perfilSheet = ss.getSheetByName("Perfil Detalhado do Creator");
    var perfilNextRow = perfilSheet.getLastRow() + 1;
    
    var perfilData = [
      "TechReview Pro", // Nome Creator
      "Decisor", // Papel na Decisão
      "Crescimento do canal e parcerias com marcas", // Objetivos Pessoais
      "Precisa de liberdade criativa e tempo para produção", // Riscos ou Objeções
      "Profissional e técnico", // Estilo de Comunicação
      "Marques Brownlee, Unbox Therapy", // Influenciadores
      "Suporte técnico e material para reviews", // Necessidades Específicas
      "Primeiro contato realizado, aguardando reunião" // Histórico de Relacionamento
    ];
    
    perfilSheet.getRange(perfilNextRow, 1, 1, perfilData.length).setValues([perfilData]);
    
    // Adiciona dados no Histórico de Campanhas
    var campSheet = ss.getSheetByName("Histórico de Campanhas");
    var campNextRow = campSheet.getLastRow() + 1;
    
    var campData = [
      "CR" + new Date().getTime().toString().slice(-6), // ID Creator
      "TechReview Pro", // Nome Creator
      new Date(), // Data Campanha
      "Review de Produto", // Tipo Campanha
      "Em Andamento", // Resultado
      "150k visualizações, 12k curtidas", // Métricas de Desempenho
      0.85, // ROI Estimado
      "Positivo, gostou da liberdade criativa" // Feedback Creator
    ];
    
    campSheet.getRange(campNextRow, 1, 1, campData.length).setValues([campData]);
    
    // Adiciona dados no Calendário Editorial
    var calSheet = ss.getSheetByName("Calendário Editorial");
    var calNextRow = calSheet.getLastRow() + 1;
    
    var calData = [
      getWeekNumber(new Date()), // Semana
      "TechReview Pro", // Creator
      "Review do Novo Smartphone XYZ", // Tema
      "Vídeo Review + Unboxing", // Formato
      "YouTube", // Plataforma
      new Date(new Date().setDate(new Date().getDate() + 10)), // Data Publicação
      "Em Produção", // Status
      "", // Link do Conteúdo
      "", // Métricas
      "Incluir unboxing e testes de desempenho" // Observações
    ];
    
    calSheet.getRange(calNextRow, 1, 1, calData.length).setValues([calData]);
    
    // Atualiza estatísticas
    atualizarEstatisticas();
    
    // Exibe mensagem de sucesso
    SpreadsheetApp.getUi().alert("Creator exemplo 'TechReview Pro' adicionado com sucesso!\n\nVocê pode usar este exemplo como referência para adicionar mais creators.");
    
  } catch (error) {
    SpreadsheetApp.getUi().alert("Erro ao criar creator exemplo: " + error.toString());
  }
}

/**
 * Gera relatório detalhado de desempenho dos creators
 */
function gerarRelatorioDesempenho() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var ui = SpreadsheetApp.getUi();
  
  try {
    // Cria nova aba para o relatório se não existir
    var reportSheet = ss.getSheetByName("Relatório de Desempenho");
    if (!reportSheet) {
      reportSheet = ss.insertSheet("Relatório de Desempenho");
    } else {
      reportSheet.clear();
    }
    
    // Configura cabeçalho do relatório
    reportSheet.getRange("A1:F1").setValues([["Relatório de Desempenho dos Creators", "", "", "", "", ""]]);
    reportSheet.getRange("A1:F1").merge();
    reportSheet.getRange("A2:F2").setValues([["Data do Relatório: " + new Date().toLocaleDateString(), "", "", "", "", ""]]);
    reportSheet.getRange("A2:F2").merge();
    
    // Adiciona seções do relatório
    reportSheet.getRange("A4:F4").setValues([["Métricas Gerais de Desempenho", "", "", "", "", ""]]);
    reportSheet.getRange("A5:B5").setValues([["Total de Creators Ativos", "=COUNTIF('Guia Principal'!L:L, \"Confirmado\")"]]);
    reportSheet.getRange("A6:B6").setValues([["Média de Engajamento", "=AVERAGE('Guia Principal'!G:G)"]]);
    reportSheet.getRange("A7:B7").setValues([["Total de Seguidores", "=SUM('Guia Principal'!F:F)"]]);
    
    // Análise por Nível de Prioridade
    reportSheet.getRange("A9:F9").setValues([["Distribuição por Nível de Prioridade", "", "", "", "", ""]]);
    reportSheet.getRange("A10:B10").setValues([["Nível 1 (Alta)", "=COUNTIF('Guia Principal'!Q:Q, \"Nível 1 (Alta)\")"]]);
    reportSheet.getRange("A11:B11").setValues([["Nível 2 (Média)", "=COUNTIF('Guia Principal'!Q:Q, \"Nível 2 (Média)\")"]]);
    reportSheet.getRange("A12:B12").setValues([["Nível 3 (Baixa)", "=COUNTIF('Guia Principal'!Q:Q, \"Nível 3 (Baixa)\")"]]);
    
    // Análise por Plataforma
    reportSheet.getRange("D10:E10").setValues([["Distribuição por Plataforma", ""]]);
    reportSheet.getRange("D11:E11").setValues([["YouTube", "=COUNTIF('Guia Principal'!E:E, \"YouTube\")"]]);
    reportSheet.getRange("D12:E12").setValues([["Instagram", "=COUNTIF('Guia Principal'!E:E, \"Instagram\")"]]);
    reportSheet.getRange("D13:E13").setValues([["TikTok", "=COUNTIF('Guia Principal'!E:E, \"TikTok\")"]]);
    
    // Formata o relatório
    reportSheet.getRange("A1:F1").setFontSize(14).setFontWeight("bold").setHorizontalAlignment("center");
    reportSheet.getRange("A4:F4").setFontWeight("bold").setBackground("#f3f3f3");
    reportSheet.getRange("A9:F9").setFontWeight("bold").setBackground("#f3f3f3");
    reportSheet.getRange("A15:F15").setFontWeight("bold").setBackground("#f3f3f3");
    
    // Ajusta largura das colunas
    reportSheet.setColumnWidth(1, 200);
    reportSheet.setColumnWidth(2, 150);
    reportSheet.setColumnWidth(3, 50);
    reportSheet.setColumnWidth(4, 200);
    reportSheet.setColumnWidth(5, 150);
    
    // Exibe mensagem de conclusão
    ui.alert("Relatório de desempenho gerado com sucesso!");
    
  } catch(error) {
    ui.alert("Erro ao gerar relatório: " + error.toString());
  }
}

/**
 * Gera análise de tendências e insights baseados nos dados históricos
 */
function gerarAnaliseInsights() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var ui = SpreadsheetApp.getUi();
  
  try {
    // Cria nova aba para análise se não existir
    var insightSheet = ss.getSheetByName("Análise e Insights");
    if (!insightSheet) {
      insightSheet = ss.insertSheet("Análise e Insights");
    } else {
      insightSheet.clear();
    }
    
    // Configura cabeçalho
    insightSheet.getRange("A1:F1").setValues([["Análise de Tendências e Insights", "", "", "", "", ""]]);
    insightSheet.getRange("A1:F1").merge();
    insightSheet.getRange("A2:F2").setValues([["Atualizado em: " + new Date().toLocaleDateString(), "", "", "", "", ""]]);
    insightSheet.getRange("A2:F2").merge();
    
    // Análise de Crescimento
    insightSheet.getRange("A4:F4").setValues([["Análise de Crescimento", "", "", "", "", ""]]);
    insightSheet.getRange("A5:B5").setValues([["Taxa de Crescimento Mensal", "=IFERROR((COUNTIF('Guia Principal'!L:L, \"Confirmado\") / LAG(COUNTIF('Guia Principal'!L:L, \"Confirmado\"), 1) - 1), \"N/A\")"]]);
    
    // Análise de Efetividade
    insightSheet.getRange("A7:F7").setValues([["Análise de Efetividade", "", "", "", "", ""]]);
    insightSheet.getRange("A8:B8").setValues([["Taxa de Conversão", "=COUNTIF('Guia Principal'!R:R, \"Fechamento\") / COUNTIF('Guia Principal'!L:L, \"Confirmado\")"]]);
    insightSheet.getRange("A9:B9").setValues([["Média de Tempo até Fechamento", "=AVERAGE(DAYS('Guia Principal'!K:K, 'Guia Principal'!N:N))"]]);
    
    // Insights de Desempenho
    insightSheet.getRange("A11:F11").setValues([["Insights de Desempenho", "", "", "", "", ""]]);
    insightSheet.getRange("A12:B12").setValues([["Creators de Alto Desempenho", "=COUNTIFS('Qualificação dos Creators'!I:I, \"Excelente*\")"]]);
    insightSheet.getRange("A13:B13").setValues([["Oportunidades de Melhoria", "=COUNTIFS('Qualificação dos Creators'!I:I, \"Regular*\", 'Qualificação dos Creators'!I:I, \"Baixo*\")"]]);
    
    // Recomendações Automáticas
    insightSheet.getRange("A15:F15").setValues([["Recomendações", "", "", "", "", ""]]);
    insightSheet.getRange("A16:F16").setValues([["Baseado nos dados analisados, recomenda-se:", "", "", "", "", ""]]);
    insightSheet.getRange("A17:F17").setValues([["• Priorizar creators com alto potencial de conversão", "", "", "", "", ""]]);
    insightSheet.getRange("A18:F18").setValues([["• Otimizar processo de qualificação", "", "", "", "", ""]]);
    insightSheet.getRange("A19:F19").setValues([["• Focar em plataformas com melhor desempenho", "", "", "", "", ""]]);
    
    // Formata a planilha
    insightSheet.getRange("A1:F1").setFontSize(14).setFontWeight("bold").setHorizontalAlignment("center");
    insightSheet.getRange("A4:F4").setFontWeight("bold").setBackground("#f3f3f3");
    insightSheet.getRange("A7:F7").setFontWeight("bold").setBackground("#f3f3f3");
    insightSheet.getRange("A11:F11").setFontWeight("bold").setBackground("#f3f3f3");
    insightSheet.getRange("A15:F15").setFontWeight("bold").setBackground("#f3f3f3");
    
    // Ajusta largura das colunas
    insightSheet.setColumnWidth(1, 250);
    insightSheet.setColumnWidth(2, 200);
    insightSheet.setColumnWidth(3, 50);
    insightSheet.setColumnWidth(4, 200);
    insightSheet.setColumnWidth(5, 150);
    
    // Exibe mensagem de conclusão
    ui.alert("Análise e insights gerados com sucesso!");
    
  } catch(error) {
    ui.alert("Erro ao gerar análise: " + error.toString());
  }
}

/**
 * Aplica formatação específica para células baseado em regras de negócio
 */
function aplicarFormatacaoEspecifica() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var ui = SpreadsheetApp.getUi();
  
  try {
    // Formata Guia Principal
    var mainSheet = ss.getSheetByName("Guia Principal");
    if (mainSheet) {
      // Verifica se há dados antes de aplicar formatação
      var lastRow = mainSheet.getLastRow();
      var lastCol = mainSheet.getLastColumn();
      
      if (lastRow > 1) { // Se houver dados além do cabeçalho
        // Formata ID do Creator
        var idRange = mainSheet.getRange(2, 1, lastRow-1, 1);
        idRange.setFontFamily("Courier New")
               .setHorizontalAlignment("left");
        
        // Formata colunas numéricas
        var numericCols = [6, 7]; // Seguidores e Engajamento
        numericCols.forEach(function(col) {
          var range = mainSheet.getRange(2, col, lastRow-1, 1);
          range.setNumberFormat("#,##0.00");
        });
        
        // Formata datas
        var dateCols = [11, 14]; // Data Primeiro Contato e Data Reunião
        dateCols.forEach(function(col) {
          var range = mainSheet.getRange(2, col, lastRow-1, 1);
          range.setNumberFormat("dd/mm/yyyy");
        });
        
        // Aplica cores condicionais para Status
        var statusRange = mainSheet.getRange(2, 12, lastRow-1, 1); // Coluna Status Contato
        var rule1 = SpreadsheetApp.newConditionalFormatRule()
          .whenTextEqualTo("Confirmado")
          .setBackground("#b7e1cd")
          .setRanges([statusRange])
          .build();
        var rule2 = SpreadsheetApp.newConditionalFormatRule()
          .whenTextEqualTo("Pendente")
          .setBackground("#fce8b2")
          .setRanges([statusRange])
          .build();
        var rule3 = SpreadsheetApp.newConditionalFormatRule()
          .whenTextEqualTo("Cancelado")
          .setBackground("#f4c7c3")
          .setRanges([statusRange])
          .build();
        
        mainSheet.setConditionalFormatRules([rule1, rule2, rule3]);
      }
    }
    
    // Formata Qualificação dos Creators
    var qualSheet = ss.getSheetByName("Qualificação dos Creators");
    if (qualSheet) {
      var qualLastRow = qualSheet.getLastRow();
      if (qualLastRow > 1) {
        // Formata pontuações (1-5)
        var scoreRange = qualSheet.getRange(2, 2, qualLastRow-1, 6); // Colunas B-G
        scoreRange.setHorizontalAlignment("center");
        
        // Aplica cores condicionais para pontuação total
        var totalScoreRange = qualSheet.getRange(2, 8, qualLastRow-1, 1); // Coluna H
        var scoreRule1 = SpreadsheetApp.newConditionalFormatRule()
          .whenNumberGreaterThanOrEqualTo(20)
          .setBackground("#b7e1cd")
          .setRanges([totalScoreRange])
          .build();
        var scoreRule2 = SpreadsheetApp.newConditionalFormatRule()
          .whenNumberBetween(15, 19)
          .setBackground("#b6d7a8")
          .setRanges([totalScoreRange])
          .build();
        var scoreRule3 = SpreadsheetApp.newConditionalFormatRule()
          .whenNumberBetween(10, 14)
          .setBackground("#fce8b2")
          .setRanges([totalScoreRange])
          .build();
        var scoreRule4 = SpreadsheetApp.newConditionalFormatRule()
          .whenNumberLessThan(10)
          .setBackground("#f4c7c3")
          .setRanges([totalScoreRange])
          .build();
        
        qualSheet.setConditionalFormatRules([scoreRule1, scoreRule2, scoreRule3, scoreRule4]);
      }
    }
    
    // Formata Calendário Editorial
    var calSheet = ss.getSheetByName("Calendário Editorial");
    if (calSheet) {
      var calLastRow = calSheet.getLastRow();
      if (calLastRow > 1) {
        // Formata data de publicação
        var pubDateRange = calSheet.getRange(2, 6, calLastRow-1, 1); // Coluna F
        pubDateRange.setNumberFormat("dd/mm/yyyy");
        
        // Aplica cores condicionais para status
        var calStatusRange = calSheet.getRange(2, 7, calLastRow-1, 1); // Coluna G
        var calRule1 = SpreadsheetApp.newConditionalFormatRule()
          .whenTextEqualTo("Publicado")
          .setBackground("#b7e1cd")
          .setRanges([calStatusRange])
          .build();
        var calRule2 = SpreadsheetApp.newConditionalFormatRule()
          .whenTextEqualTo("Em Produção")
          .setBackground("#fce8b2")
          .setRanges([calStatusRange])
          .build();
        var calRule3 = SpreadsheetApp.newConditionalFormatRule()
          .whenTextEqualTo("Atrasado")
          .setBackground("#f4c7c3")
          .setRanges([calStatusRange])
          .build();
        
        calSheet.setConditionalFormatRules([calRule1, calRule2, calRule3]);
      }
    }
    
    ui.alert("Formatação aplicada com sucesso!");
    
  } catch(error) {
    ui.alert("Erro ao aplicar formatação: " + error.toString());
  }
}

/**
 * Função utilitária para obter o número da semana
 * @param {Date} date Data para calcular o número da semana
 * @return {number} Número da semana no ano
 */
function getWeekNumber(date) {
  var d = new Date(Date.UTC(date.getFullYear(), date.getMonth(), date.getDate()));
  var dayNum = d.getUTCDay() || 7;
  d.setUTCDate(d.getUTCDate() + 4 - dayNum);
  var yearStart = new Date(Date.UTC(d.getUTCFullYear(),0,1));
  return Math.ceil((((d - yearStart) / 86400000) + 1)/7);
}

/**
 * Função utilitária para validar endereço de email
 * @param {string} email Endereço de email para validar
 * @return {boolean} Verdadeiro se o email é válido
 */
function validarEmail(email) {
  var regex = /^[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}$/;
  return regex.test(email);
}

/**
 * Função utilitária para validar número de telefone
 * @param {string} telefone Número de telefone para validar
 * @return {boolean} Verdadeiro se o telefone é válido
 */
function validarTelefone(telefone) {
  var regex = /^\+?[\d\s-()]{10,}$/;
  return regex.test(telefone);
}

/**
 * Função utilitária para formatar números grandes
 * @param {number} numero Número para formatar
 * @return {string} Número formatado com sufixo K ou M
 */
function formatarNumeroGrande(numero) {
  if (numero >= 1000000) {
    return (numero / 1000000).toFixed(1) + "M";
  } else if (numero >= 1000) {
    return (numero / 1000).toFixed(1) + "K";
  }
  return numero.toString();
}

/**
 * Valida os dados de entrada de um novo creator
 * @param {Object} dados Objeto com os dados do creator
 * @return {Object} Objeto com resultado da validação e mensagens de erro
 */
function validarDadosCreator(dados) {
  var erros = [];
  
  // Valida campos obrigatórios
  if (!dados.nome) erros.push("Nome do creator é obrigatório");
  if (!dados.categoria) erros.push("Categoria é obrigatória");
  if (!dados.plataforma) erros.push("Plataforma principal é obrigatória");
  if (!dados.seguidores) erros.push("Número de seguidores é obrigatório");
  
  // Valida formato dos dados
  if (dados.email && !validarEmail(dados.email)) {
    erros.push("Formato de email inválido");
  }
  
  if (dados.telefone && !validarTelefone(dados.telefone)) {
    erros.push("Formato de telefone inválido");
  }
  
  if (dados.seguidores && isNaN(dados.seguidores)) {
    erros.push("Número de seguidores deve ser um número");
  }
  
  if (dados.engajamento && (isNaN(dados.engajamento) || dados.engajamento < 0 || dados.engajamento > 100)) {
    erros.push("Taxa de engajamento deve ser um número entre 0 e 100");
  }
  
  return {
    valido: erros.length === 0,
    erros: erros
  };
}

/**
 * Calcula métricas de desempenho para um creator
 * @param {Object} dados Objeto com os dados do creator
 * @return {Object} Objeto com as métricas calculadas
 */
function calcularMetricasDesempenho(dados) {
  var metricas = {
    pontuacaoTotal: 0,
    classificacao: "",
    potencialROI: 0
  };
  
  // Calcula pontuação baseada em vários fatores
  var pontuacaoSeguidores = calcularPontuacaoSeguidores(dados.seguidores);
  var pontuacaoEngajamento = calcularPontuacaoEngajamento(dados.engajamento);
  var pontuacaoConversao = dados.taxaConversao ? calcularPontuacaoConversao(dados.taxaConversao) : 0;
  
  metricas.pontuacaoTotal = pontuacaoSeguidores + pontuacaoEngajamento + pontuacaoConversao;
  
  // Define classificação baseada na pontuação total
  if (metricas.pontuacaoTotal >= 20) {
    metricas.classificacao = "Excelente (20-25)";
  } else if (metricas.pontuacaoTotal >= 15) {
    metricas.classificacao = "Bom (15-19)";
  } else if (metricas.pontuacaoTotal >= 10) {
    metricas.classificacao = "Regular (10-14)";
  } else {
    metricas.classificacao = "Baixo Potencial (<10)";
  }
  
  // Calcula potencial ROI baseado nas métricas
  metricas.potencialROI = calcularPotencialROI(dados);
  
  return metricas;
}

/**
 * Calcula pontuação baseada no número de seguidores
 * @param {number} seguidores Número de seguidores
 * @return {number} Pontuação (1-5)
 */
function calcularPontuacaoSeguidores(seguidores) {
  if (seguidores >= 1000000) return 5;
  if (seguidores >= 500000) return 4;
  if (seguidores >= 100000) return 3;
  if (seguidores >= 50000) return 2;
  return 1;
}

/**
 * Calcula pontuação baseada na taxa de engajamento
 * @param {number} engajamento Taxa de engajamento em porcentagem
 * @return {number} Pontuação (1-5)
 */
function calcularPontuacaoEngajamento(engajamento) {
  if (engajamento >= 10) return 5;
  if (engajamento >= 7) return 4;
  if (engajamento >= 5) return 3;
  if (engajamento >= 3) return 2;
  return 1;
}

/**
 * Calcula pontuação baseada na taxa de conversão histórica
 * @param {number} taxaConversao Taxa de conversão em porcentagem
 * @return {number} Pontuação (1-5)
 */
function calcularPontuacaoConversao(taxaConversao) {
  if (taxaConversao >= 5) return 5;
  if (taxaConversao >= 3) return 4;
  if (taxaConversao >= 2) return 3;
  if (taxaConversao >= 1) return 2;
  return 1;
}

/**
 * Calcula potencial ROI baseado nos dados do creator
 * @param {Object} dados Objeto com os dados do creator
 * @return {number} Potencial ROI estimado
 */
function calcularPotencialROI(dados) {
  var baseROI = 1.0;
  
  // Ajusta baseado no número de seguidores
  if (dados.seguidores >= 1000000) baseROI *= 1.5;
  else if (dados.seguidores >= 500000) baseROI *= 1.3;
  else if (dados.seguidores >= 100000) baseROI *= 1.2;
  
  // Ajusta baseado no engajamento
  if (dados.engajamento >= 10) baseROI *= 1.5;
  else if (dados.engajamento >= 7) baseROI *= 1.3;
  else if (dados.engajamento >= 5) baseROI *= 1.2;
  
  // Ajusta baseado na taxa de conversão histórica
  if (dados.taxaConversao) {
    if (dados.taxaConversao >= 5) baseROI *= 1.5;
    else if (dados.taxaConversao >= 3) baseROI *= 1.3;
    else if (dados.taxaConversao >= 2) baseROI *= 1.2;
  }
  
  return baseROI;
}

/**
 * Adiciona um novo creator à planilha principal
 * @param {Object} dados Dados do creator para adicionar
 */
function adicionarCreator(dados) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var ui = SpreadsheetApp.getUi();
  
  try {
    // Valida os dados primeiro
    var validacao = validarDadosCreator(dados);
    if (!validacao.valido) {
      ui.alert("Erro ao adicionar creator:\n\n" + validacao.erros.join("\n"));
      return;
    }
    
    // Gera ID único para o creator
    var idCreator = "CR" + new Date().getTime().toString().slice(-6);
    
    // Adiciona à Guia Principal
    var mainSheet = ss.getSheetByName("Guia Principal");
    var nextRow = mainSheet.getLastRow() + 1;
    
    var rowData = [
      idCreator,
      dados.nome,
      dados.categoria,
      dados.nicho,
      dados.plataforma,
      dados.seguidores,
      dados.engajamento,
      dados.email,
      dados.telefone,
      dados.linkPerfil,
      new Date(), // Data Primeiro Contato
      "Pendente", // Status Contato
      "Qualificar", // Próxima Ação
      "", // Data Reunião
      "", // Horário Reunião
      Session.getActiveUser().getEmail(), // Responsável Interno
      "Nível 3 (Baixa)", // Prioridade inicial
      "Prospecção", // Etapa inicial
      "Básico", // Nível Personalização inicial
      dados.observacoes || "" // Observações
    ];
    
    mainSheet.getRange(nextRow, 1, 1, rowData.length).setValues([rowData]);
    
    // Adiciona à Qualificação dos Creators
    var qualSheet = ss.getSheetByName("Qualificação dos Creators");
    var qualNextRow = qualSheet.getLastRow() + 1;
    
    // Calcula métricas iniciais
    var metricas = calcularMetricasDesempenho(dados);
    
    var qualData = [
      dados.nome,
      0, // Potencial Alcance
      0, // Relevância para Audiência
      0, // Compatibilidade com Marca
      0, // Qualidade de Conteúdo
      0, // Engajamento da Audiência
      0, // Histórico de Conversão
      "=SUM(B" + qualNextRow + ":G" + qualNextRow + ")", // Pontuação Total
      "=IF(H" + qualNextRow + ">=20, \"Excelente (20-25)\", IF(H" + qualNextRow + ">=15, \"Bom (15-19)\", IF(H" + qualNextRow + ">=10, \"Regular (10-14)\", \"Baixo Potencial (<10)\")))", // Status Qualificação
      "A Definir", // Segmento
      "Nível 3 (Baixa)" // Nível de Prioridade inicial
    ];
    
    qualSheet.getRange(qualNextRow, 1, 1, qualData.length).setValues([qualData]);
    
    // Adiciona ao Perfil Detalhado
    var perfilSheet = ss.getSheetByName("Perfil Detalhado do Creator");
    var perfilNextRow = perfilSheet.getLastRow() + 1;
    
    var perfilData = [
      dados.nome,
      "", // Papel na Decisão
      "", // Objetivos Pessoais
      "", // Riscos ou Objeções
      "", // Estilo de Comunicação
      "", // Influenciadores
      "", // Necessidades Específicas
      "Primeiro contato pendente" // Histórico de Relacionamento
    ];
    
    perfilSheet.getRange(perfilNextRow, 1, 1, perfilData.length).setValues([perfilData]);
    
    // Atualiza estatísticas
    atualizarEstatisticas();
    
    ui.alert("Creator adicionado com sucesso!");
    
  } catch(error) {
    ui.alert("Erro ao adicionar creator: " + error.toString());
  }
}

/**
 * Atualiza os dados de um creator existente
 * @param {string} idCreator ID do creator para atualizar
 * @param {Object} dados Novos dados do creator
 */
function atualizarCreator(idCreator, dados) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var ui = SpreadsheetApp.getUi();
  
  try {
    // Valida os dados primeiro
    var validacao = validarDadosCreator(dados);
    if (!validacao.valido) {
      ui.alert("Erro ao atualizar creator:\n\n" + validacao.erros.join("\n"));
      return;
    }
    
    // Atualiza Guia Principal
    var mainSheet = ss.getSheetByName("Guia Principal");
    var mainData = mainSheet.getDataRange().getValues();
    var mainRow = -1;
    
    // Encontra a linha do creator
    for (var i = 0; i < mainData.length; i++) {
      if (mainData[i][0] === idCreator) {
        mainRow = i + 1;
        break;
      }
    }
    
    if (mainRow === -1) {
      ui.alert("Creator não encontrado!");
      return;
    }
    
    // Atualiza dados na Guia Principal
    var updateData = [
      idCreator,
      dados.nome,
      dados.categoria,
      dados.nicho,
      dados.plataforma,
      dados.seguidores,
      dados.engajamento,
      dados.email,
      dados.telefone,
      dados.linkPerfil,
      mainData[mainRow-1][10], // Mantém Data Primeiro Contato
      dados.statusContato || mainData[mainRow-1][11],
      dados.proximaAcao || mainData[mainRow-1][12],
      dados.dataReuniao || mainData[mainRow-1][13],
      dados.horarioReuniao || mainData[mainRow-1][14],
      mainData[mainRow-1][15], // Mantém Responsável Interno
      dados.prioridade || mainData[mainRow-1][16],
      dados.etapa || mainData[mainRow-1][17],
      dados.nivelPersonalizacao || mainData[mainRow-1][18],
      dados.observacoes || mainData[mainRow-1][19]
    ];
    
    mainSheet.getRange(mainRow, 1, 1, updateData.length).setValues([updateData]);
    
    // Atualiza outras abas relacionadas
    atualizarQualificacaoCreator(dados.nome, dados);
    atualizarPerfilCreator(dados.nome, dados);
    
    // Atualiza estatísticas
    atualizarEstatisticas();
    
    ui.alert("Creator atualizado com sucesso!");
    
  } catch(error) {
    ui.alert("Erro ao atualizar creator: " + error.toString());
  }
}

/**
 * Atualiza a qualificação de um creator
 * @param {string} nomeCreator Nome do creator
 * @param {Object} dados Dados atualizados
 */
function atualizarQualificacaoCreator(nomeCreator, dados) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var qualSheet = ss.getSheetByName("Qualificação dos Creators");
  var qualData = qualSheet.getDataRange().getValues();
  var qualRow = -1;
  
  // Encontra a linha do creator
  for (var i = 0; i < qualData.length; i++) {
    if (qualData[i][0] === nomeCreator) {
      qualRow = i + 1;
      break;
    }
  }
  
  if (qualRow !== -1) {
    // Atualiza pontuações se fornecidas
    var updateData = [
      nomeCreator,
      dados.potencialAlcance || qualData[qualRow-1][1],
      dados.relevanciaAudiencia || qualData[qualRow-1][2],
      dados.compatibilidadeMarca || qualData[qualRow-1][3],
      dados.qualidadeConteudo || qualData[qualRow-1][4],
      dados.engajamentoAudiencia || qualData[qualRow-1][5],
      dados.historicoConversao || qualData[qualRow-1][6],
      "=SUM(B" + (qualSheet.getLastRow()) + ":G" + (qualSheet.getLastRow()) + ")", // Pontuação Total
      "=IF(H" + (qualSheet.getLastRow()) + ">=20, \"Excelente (20-25)\", IF(H" + (qualSheet.getLastRow()) + ">=15, \"Bom (15-19)\", IF(H" + (qualSheet.getLastRow()) + ">=10, \"Regular (10-14)\", \"Baixo Potencial (<10)\")))", // Status Qualificação
      dados.segmento || qualData[qualRow-1][8], // Segmento
      dados.nivelPrioridade || qualData[qualRow-1][9] // Nível de Prioridade
    ];
    
    qualSheet.getRange(qualRow, 1, 1, updateData.length).setValues([updateData]);
  }
}

/**
 * Gerencia eventos do calendário de reuniões
 */
function gerenciarEventosCalendario() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var ui = SpreadsheetApp.getUi();
  
  try {
    var calSheet = ss.getSheetByName("Calendário de Reuniões");
    if (!calSheet) {
      ui.alert("Guia de Calendário não encontrada!");
      return;
    }
    
    var data = calSheet.getDataRange().getValues();
    var calendar = CalendarApp.getDefaultCalendar();
    var eventosAtualizados = 0;
    var eventosNovos = 0;
    
    // Começa da linha 1 (pula o cabeçalho)
    for (var i = 1; i < data.length; i++) {
      var dataReuniao = data[i][0]; // Data
      var horario = data[i][1]; // Horário
      var creator = data[i][2]; // Nome Creator
      var plataforma = data[i][3]; // Plataforma
      var link = data[i][4]; // Link
      var status = data[i][5]; // Status
      var idEvento = data[i][6]; // ID do Evento
      
      // Verifica se tem data e hora válidas
      if (dataReuniao && horario) {
        // Cria uma data completa concatenando data e hora
        var dataCompleta = new Date(dataReuniao);
        var horaParts = horario.split(":");
        if (horaParts.length >= 2) {
          dataCompleta.setHours(parseInt(horaParts[0]), parseInt(horaParts[1]));
          
          // Cria evento de 1 hora
          var endTime = new Date(dataCompleta.getTime() + 60 * 60 * 1000);
          
          // Descrição do evento
          var desc = "Reunião com creator: " + creator;
          if (plataforma) desc += "\nPlataforma: " + plataforma;
          if (link) desc += "\nLink: " + link;
          
          // Verifica se já existe um evento
          if (idEvento) {
            try {
              var evento = calendar.getEventById(idEvento);
              if (evento) {
                // Atualiza evento existente
                evento.setTime(dataCompleta, endTime);
                evento.setDescription(desc);
                eventosAtualizados++;
              } else {
                // Cria novo evento se o anterior não existir
                var novoEvento = calendar.createEvent(
                  "Reunião: " + creator,
                  dataCompleta,
                  endTime,
                  {description: desc}
                );
                calSheet.getRange(i + 1, 7).setValue(novoEvento.getId());
                eventosNovos++;
              }
            } catch (e) {
              // Se houver erro ao acessar o evento, cria um novo
              var novoEvento = calendar.createEvent(
                "Reunião: " + creator,
                dataCompleta,
                endTime,
                {description: desc}
              );
              calSheet.getRange(i + 1, 7).setValue(novoEvento.getId());
              eventosNovos++;
            }
          } else {
            // Cria novo evento
            var novoEvento = calendar.createEvent(
              "Reunião: " + creator,
              dataCompleta,
              endTime,
              {description: desc}
            );
            calSheet.getRange(i + 1, 7).setValue(novoEvento.getId());
            eventosNovos++;
          }
        }
      }
    }
    
    ui.alert("Calendário atualizado com sucesso!\n" +
             eventosAtualizados + " eventos atualizados\n" +
             eventosNovos + " novos eventos criados");
             
  } catch(error) {
    ui.alert("Erro ao gerenciar eventos do calendário: " + error.toString());
  }
}

/**
 * Sincroniza eventos do calendário com a planilha
 */
function sincronizarCalendario() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var ui = SpreadsheetApp.getUi();
  
  try {
    var calSheet = ss.getSheetByName("Calendário de Reuniões");
    if (!calSheet) {
      ui.alert("Guia de Calendário não encontrada!");
      return;
    }
    
    var calendar = CalendarApp.getDefaultCalendar();
    var hoje = new Date();
    var tresMesesDepois = new Date(hoje.getTime() + (90 * 24 * 60 * 60 * 1000));
    
    // Busca eventos no calendário
    var eventos = calendar.getEvents(hoje, tresMesesDepois);
    var eventosReuniao = eventos.filter(function(evento) {
      return evento.getTitle().indexOf("Reunião:") === 0;
    });
  } catch(error) {
    ui.alert("Erro ao sincronizar calendário: " + error.toString());
  }
}

/**
 * Limpa dados antigos e desnecessários da planilha
 */
function limparDadosAntigos() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var ui = SpreadsheetApp.getUi();
  
  // Confirma com o usuário antes de prosseguir
  var response = ui.alert(
    'Limpar Dados Antigos',
    'Isso removerá dados antigos e desnecessários da planilha. Deseja continuar?',
    ui.ButtonSet.YES_NO
  );
  
  if (response !== ui.Button.YES) {
    return;
  }
  
  try {
    // Limpa reuniões antigas do calendário
    var calSheet = ss.getSheetByName("Calendário de Reuniões");
    if (calSheet) {
      var calData = calSheet.getDataRange().getValues();
      var hoje = new Date();
      var linhasParaRemover = [];
      
      // Identifica reuniões antigas (mais de 3 meses)
      for (var i = 1; i < calData.length; i++) {
        var dataReuniao = calData[i][0];
        if (dataReuniao && dataReuniao instanceof Date) {
          var diffMeses = (hoje.getTime() - dataReuniao.getTime()) / (30 * 24 * 60 * 60 * 1000);
          if (diffMeses > 3) {
            linhasParaRemover.push(i + 1);
          }
        }
      }
      
      // Remove linhas de baixo para cima para não afetar os índices
      for (var i = linhasParaRemover.length - 1; i >= 0; i--) {
        calSheet.deleteRow(linhasParaRemover[i]);
      }
    }
    
    // Limpa histórico de campanhas antigas
    var campSheet = ss.getSheetByName("Histórico de Campanhas");
    if (campSheet) {
      var campData = campSheet.getDataRange().getValues();
      var linhasParaRemover = [];
      
      // Identifica campanhas antigas (mais de 6 meses)
      for (var i = 1; i < campData.length; i++) {
        var dataCampanha = campData[i][2];
        if (dataCampanha && dataCampanha instanceof Date) {
          var diffMeses = (hoje.getTime() - dataCampanha.getTime()) / (30 * 24 * 60 * 60 * 1000);
          if (diffMeses > 6) {
            linhasParaRemover.push(i + 1);
          }
        }
      }
      
      // Remove linhas de baixo para cima para não afetar os índices
      for (var i = linhasParaRemover.length - 1; i >= 0; i--) {
        campSheet.deleteRow(linhasParaRemover[i]);
      }
    }
    
    ui.alert("Dados antigos limpos com sucesso!");
    
  } catch(error) {
    ui.alert("Erro ao limpar dados antigos: " + error.toString());
  }
}

/**
 * Configura gatilhos automáticos para as funções principais
 */
function configurarGatilhosAutomaticos() {
  var ui = SpreadsheetApp.getUi();
  
  try {
    // Remove gatilhos existentes
    var gatilhos = ScriptApp.getProjectTriggers();
    for (var i = 0; i < gatilhos.length; i++) {
      ScriptApp.deleteTrigger(gatilhos[i]);
    }
    
    // Configura gatilho para atualização diária de estatísticas
    ScriptApp.newTrigger('atualizarEstatisticas')
      .timeBased()
      .everyDays(1)
      .atHour(1)
      .create();
    
    // Configura gatilho para sincronização do calendário a cada 6 horas
    ScriptApp.newTrigger('sincronizarCalendario')
      .timeBased()
      .everyHours(6)
      .create();
    
    // Configura gatilho para limpeza mensal de dados antigos
    ScriptApp.newTrigger('limparDadosAntigos')
      .timeBased()
      .onMonthDay(1)
      .atHour(2)
      .create();
    
    ui.alert("Gatilhos automáticos configurados com sucesso!");
    
  } catch(error) {
    ui.alert("Erro ao configurar gatilhos: " + error.toString());
  }
}

/**
 * Verifica e notifica sobre tarefas pendentes
 */
function verificarTarefasPendentes() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var ui = SpreadsheetApp.getUi();
  
  try {
    var mainSheet = ss.getSheetByName("Guia Principal");
    if (!mainSheet) {
      ui.alert("Guia Principal não encontrada!");
      return;
    }
    
    var data = mainSheet.getDataRange().getValues();
    var tarefasPendentes = [];
    var hoje = new Date();
    
    // Verifica cada linha por tarefas pendentes
    for (var i = 1; i < data.length; i++) {
      var proximaAcao = data[i][12]; // Próxima Ação
      var dataReuniao = data[i][13]; // Data Reunião
      var creator = data[i][1]; // Nome Creator
      
      if (proximaAcao && proximaAcao !== "Nenhuma") {
        if (dataReuniao && dataReuniao instanceof Date) {
          // Verifica se a reunião está próxima
          var diffDias = (dataReuniao.getTime() - hoje.getTime()) / (24 * 60 * 60 * 1000);
          if (diffDias <= 2 && diffDias > 0) {
            tarefasPendentes.push("Reunião em " + Math.ceil(diffDias) + " dias com " + creator);
          }
        } else {
          tarefasPendentes.push("Ação pendente para " + creator + ": " + proximaAcao);
        }
      }
    }
    
    // Verifica qualificações pendentes
    var qualSheet = ss.getSheetByName("Qualificação dos Creators");
    if (qualSheet) {
      var qualData = qualSheet.getDataRange().getValues();
      for (var i = 1; i < qualData.length; i++) {
        var pontuacaoTotal = qualData[i][7]; // Pontuação Total
        if (pontuacaoTotal === 0) {
          tarefasPendentes.push("Qualificação pendente para " + qualData[i][0]);
        }
      }
    }
    
    // Exibe notificação com tarefas pendentes
    if (tarefasPendentes.length > 0) {
      ui.alert(
        "Tarefas Pendentes",
        "Você tem " + tarefasPendentes.length + " tarefas pendentes:\n\n" +
        tarefasPendentes.join("\n"),
        ui.ButtonSet.OK
      );
    } else {
      ui.alert("Não há tarefas pendentes no momento!");
    }
    
  } catch(error) {
    ui.alert("Erro ao verificar tarefas pendentes: " + error.toString());
  }
}

/**
 * Gera backup dos dados importantes
 */
function gerarBackup() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var ui = SpreadsheetApp.getUi();
  
  try {
    // Cria nova planilha para backup
    var backupSS = SpreadsheetApp.create("Backup_Creators_" + new Date().toISOString().slice(0,10));
    
    // Lista de abas para fazer backup
    var abasParaBackup = [
      "Guia Principal",
      "Qualificação dos Creators",
      "Perfil Detalhado do Creator",
      "Histórico de Campanhas",
      "Calendário de Reuniões"
    ];
    
    // Copia cada aba para o backup
    abasParaBackup.forEach(function(nomeAba) {
      var sourceSheet = ss.getSheetByName(nomeAba);
      if (sourceSheet) {
        // Copia dados e formatação
        var targetSheet = backupSS.insertSheet(nomeAba);
        var range = sourceSheet.getDataRange();
        var data = range.getValues();
        var formats = range.getNumberFormats();
        var backgrounds = range.getBackgrounds();
        var fontColors = range.getFontColors();
        var fontFamilies = range.getFontFamilies();
        var fontSizes = range.getFontSizes();
        var fontLines = range.getFontLines();
        var fontWeights = range.getFontWeights();
        var horizontalAlignments = range.getHorizontalAlignments();
        var verticalAlignments = range.getVerticalAlignments();
        
        // Aplica dados e formatação
        var targetRange = targetSheet.getRange(1, 1, data.length, data[0].length);
        targetRange.setValues(data);
        targetRange.setNumberFormats(formats);
        targetRange.setBackgrounds(backgrounds);
        targetRange.setFontColors(fontColors);
        targetRange.setFontFamilies(fontFamilies);
        targetRange.setFontSizes(fontSizes);
        targetRange.setFontLines(fontLines);
        targetRange.setFontWeights(fontWeights);
        targetRange.setHorizontalAlignments(horizontalAlignments);
        targetRange.setVerticalAlignments(verticalAlignments);
        
        // Ajusta largura das colunas
        for (var i = 1; i <= data[0].length; i++) {
          targetSheet.setColumnWidth(i, sourceSheet.getColumnWidth(i));
        }
      }
    });
    
  } catch(error) {
    ui.alert("Erro ao gerar backup: " + error.toString());
  }
}

/**
 * Função auxiliar para formatar data no padrão brasileiro
 * @param {Date} data Data para formatar
 * @return {string} Data formatada
 */
function formatarData(data) {
  if (!data || !(data instanceof Date)) return "";
  return data.getDate().toString().padStart(2, '0') + '/' +
         (data.getMonth() + 1).toString().padStart(2, '0') + '/' +
         data.getFullYear();
}

/**
 * Função auxiliar para formatar moeda no padrão brasileiro
 * @param {number} valor Valor para formatar
 * @return {string} Valor formatado
 */
function formatarMoeda(valor) {
  if (isNaN(valor)) return "R$ 0,00";
  return "R$ " + valor.toFixed(2).replace('.', ',').replace(/\B(?=(\d{3})+(?!\d))/g, ".");
}

/**
 * Função auxiliar para formatar porcentagem
 * @param {number} valor Valor para formatar (0-100)
 * @return {string} Valor formatado
 */
function formatarPorcentagem(valor) {
  if (isNaN(valor)) return "0%";
  return valor.toFixed(1).replace('.', ',') + "%";
}

/**
 * Função auxiliar para gerar ID único
 * @return {string} ID único gerado
 */
function gerarIdUnico() {
  return "CR" + new Date().getTime().toString().slice(-6);
}

/**
 * Função auxiliar para calcular diferença em dias entre duas datas
 * @param {Date} data1 Primeira data
 * @param {Date} data2 Segunda data
 * @return {number} Diferença em dias
 */
function calcularDiferencaDias(data1, data2) {
  return Math.round((data2.getTime() - data1.getTime()) / (24 * 60 * 60 * 1000));
}

/**
 * Função auxiliar para validar dados numéricos
 * @param {any} valor Valor para validar
 * @param {number} min Valor mínimo permitido
 * @param {number} max Valor máximo permitido
 * @return {boolean} Verdadeiro se o valor é válido
 */
function validarNumero(valor, min, max) {
  var num = parseFloat(valor);
  return !isNaN(num) && num >= min && num <= max;
}

/**
 * Função auxiliar para limpar formatação de texto
 * @param {string} texto Texto para limpar
 * @return {string} Texto limpo
 */
function limparTexto(texto) {
  return texto.trim().replace(/\s+/g, ' ');
}

/**
 * Função auxiliar para validar nome do creator
 * @param {string} nome Nome para validar
 * @return {boolean} Verdadeiro se o nome é válido
 */
function validarNomeCreator(nome) {
  return nome && nome.trim().length >= 3 && nome.trim().length <= 100;
}

// Fim do arquivo