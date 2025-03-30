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
 * Cria um menu personalizado na interface com mais opções
 */
function onOpen() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu('Gestão de Creators')
    .addItem('Criar/Redefinir Planilha', 'criarPlanilhaCompleta')
    .addSeparator()
    .addItem('Atualizar Estatísticas', 'atualizarEstatisticas')
    .addItem('Gerar Relatório', 'gerarRelatorio')
    .addSeparator()
    .addItem('Exportar Calendário de Reuniões', 'exportarCalendario')
    .addSeparator()
    .addItem('Configurar Dashboard', 'configurarDashboard')
    .addItem('Inicializar Página de Instruções', 'inicializarPaginaInstrucoes')
    .addItem('Inicializar Sumário Executivo', 'inicializarSumarioExecutivo')
    .addToUi();
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
  // ===== SEÇÃO 5: PRINCIPAIS INSIGHTS E RECOMENDAÇÕES =====
  var nextRow = Math.max(campRow + 2, nextRow + 12);
  
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
 * Funções auxiliares para o relatório
 */
function getWeekNumber(d) {
  d = new Date(Date.UTC(d.getFullYear(), d.getMonth(), d.getDate()));
  d.setUTCDate(d.getUTCDate() + 4 - (d.getUTCDay()||7));
  var yearStart = new Date(Date.UTC(d.getUTCFullYear(),0,1));
  var weekNo = Math.ceil(( ( (d - yearStart) / 86400000) + 1)/7);
  return weekNo;
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
 * Configura o dashboard com gráficos e métricas avançadas baseadas no ABM Template
 */
function configurarDashboard() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var dashSheet = ss.getSheetByName("Dashboard");
  
  // Limpa o dashboard existente
  dashSheet.clear();
  
  // Configuração básica da planilha
  dashSheet.setColumnWidth(1, 200);
  dashSheet.setColumnWidth(2, 150);
  dashSheet.setColumnWidth(3, 150);
  dashSheet.setColumnWidth(4, 150);
  dashSheet.setColumnWidth(5, 150);
  
  // Adiciona título e subtítulo
  dashSheet.getRange("A1:E1").merge();
  dashSheet.getRange("A1").setValue("DASHBOARD DE GESTÃO DE CREATORS");
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
  dashSheet.getRange("A5").setValue("KPIs PRINCIPAIS");
  dashSheet.getRange("A5").setFontWeight("bold").setHorizontalAlignment("center");
  dashSheet.getRange("A5:E5").setBackground("#e3f2fd");
  
  // Layout de 2x3 para KPIs
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
 * Script para gerenciamento de creators
 * Funcionalidades: criação de planilha, automações, formatação, validações e relatórios
 */

// Constantes globais
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

const PRIORIDADE_OPTIONS = [
  "Tier 1 (Alta)",
  "Tier 2 (Média)", 
  "Tier 3 (Baixa)",
  "Backlog"
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

const CHECKPOINT_OPTIONS = [
  "Q1 2023",
  "Q2 2023",
  "Q3 2023",
  "Q4 2023",
  "Q1 2024",
  "Q2 2024",
  "Q3 2024",
  "Q4 2024"
];

/**
 * Função principal para criar/redefinir todas as planilhas
 */
function criarPlanilhaCompleta() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();

  var guias = [
    // Página de Instruções e Apresentação
    {nome: "Página de Instruções", colunas: ["Instruções de Uso da Planilha"]},
    {nome: "Sumário Executivo", colunas: ["Resumo do Programa de Creators", "KPIs Principais", "Período Atual", "Status Geral"]},
    {nome: "Jornada do Seguidor", colunas: ["Mapeamento da Jornada de Consumo"]},
    
    // Gestão de Creators (núcleo)
    {nome: "Guia Principal", colunas: ["ID", "Nome", "Categoria", "Nicho Específico", "Plataforma Principal", "Seguidores", "Engajamento (%)", "Email", "WhatsApp", "Link Perfil", "Data Primeiro Contato", "Status Contato", "Próxima Ação", "Data Reunião", "Horário Reunião", "Responsável Interno", "Prioridade", "Etapa Jornada Compra", "Nível Personalização", "Notas"]},
    {nome: "Calendário de Reuniões", colunas: ["Data", "Horário", "Nome Creator", "Plataforma (Zoom/Meet)", "Link da Reunião", "Status da Reunião"]},
    {nome: "Histórico de Campanhas", colunas: ["ID Creator", "Nome Creator", "Data Campanha", "Tipo Campanha", "Resultado", "Métricas de Desempenho", "ROI Estimado", "Feedback Creator"]},
    
    // Análise e Qualificação (inspirado no ABM Template)
    {nome: "Qualificação dos Creators", colunas: ["Nome Creator", "Potencial Alcance (1-5)", "Relevância para Audiência (1-5)", "Compatibilidade com Marca (1-5)", "Qualidade de Conteúdo (1-5)", "Engajamento da Audiência (1-5)", "Histórico de Conversão (1-5)", "Pontuação Total", "Status Qualificação", "Segmento", "Tier de Prioridade"]},
    {nome: "ICP de Creators", colunas: ["Critério", "Perfil Ideal (ICP)", "Creator 1", "Creator 2", "Creator 3", "Creator 4", "Creator 5", "Creator 6", "Creator 7", "Creator 8"]},
    {nome: "Perfil Detalhado do Creator", colunas: ["Nome Creator", "Papel Decisão", "Objetivos Pessoais", "Riscos ou Objeções", "Estilo de Comunicação", "Influenciadores", "Necessidades Específicas", "Histórico de Relacionamento"]},
    {nome: "Análise Competitiva", colunas: ["Creator", "Trabalha com Concorrente 1", "Trabalha com Concorrente 2", "Trabalha com Concorrente 3", "Nossa Proposta Valor", "Diferencial Competitivo", "Estratégia de Abordagem"]},
    {nome: "Análise de Interesses", colunas: ["Creators Envolvidos", "Interesses Compartilhados", "Pontos de Divergência", "Potencial Colaboração", "Ações Recomendadas"]},
    
    // Planejamento Estratégico
    {nome: "Alinhamento Estratégico", colunas: ["Objetivo de Negócio", "Estratégia com Creators", "KPIs", "Métricas de Sucesso", "Tática", "Responsável", "Status"]},
    {nome: "Alocação de Recursos", colunas: ["Categoria", "Orçamento Estimado", "Orçamento Realizado", "Variação", "Observações"]},
    {nome: "Playbook de Engajamento", colunas: ["Etapa Jornada", "Ações de Marketing", "Ações de Vendas", "Canais Específicos", "Responsável pela Ação"]},
    
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
  inicializarAnaliseCompetitiva();
  inicializarJornadaSeguidor();
  inicializarMatrizPersonalizacao();
  inicializarPlaybookEngajamento(); // Nova chamada para inicializar o playbook
  inicializarAnaliseInteresses(); // Nova chamada para inicializar a análise de interesses
  
  // Configura o dashboard
  configurarDashboard();

  // Atualiza estatísticas
  atualizarEstatisticas();

  // Exibe mensagem de confirmação
  SpreadsheetApp.getUi().alert("Planilha criada com sucesso!\n\nUtilize o menu 'Gestão de Creators' para acessar todas as funcionalidades.");
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
  
  // Atualiza o dashboard
  var dashSheet = ss.getSheetByName("Dashboard");
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
      if(col >= 2 && col <= 7 && row > 4){ // Colunas de pontuação (B-G)
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
            tierCell.setValue("Tier 1 (Alta)");
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
            tierCell.setValue("Tier 2 (Média)");
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
            tierCell.setValue("Tier 3 (Baixa)");
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
            tierCell.setValue("Backlog");
          }
        }
      }
      
      // Atualiza pontuação de personalização quando qualquer critério é alterado
      if(col >= 12 && col <= 16 && row > 4){ // Colunas de personalização (L-P)
        var personalizacaoCell = sheet.getRange(row, 17); // Coluna Q = Pontuação Personalização
        
        // Verifica se a fórmula de soma está presente
        if(!personalizacaoCell.getFormula()){
          personalizacaoCell.setFormula("=SUM(L" + row + ":P" + row + ")");
        }
        
        // Aplica formatação condicional baseada na pontuação
        var pontuacao = personalizacaoCell.getValue();
        if(pontuacao >= 20){
          sheet.getRange(row, 17).setBackground("#c8e6c9"); // Verde claro
        } else if(pontuacao >= 15){
          sheet.getRange(row, 17).setBackground("#fff9c4"); // Amarelo claro
        } else {
          sheet.getRange(row, 17).setBackground("#ffcdd2"); // Vermelho claro
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
 * Aplicar formatação específica para cada guia
 */
function aplicarFormatacaoEspecifica(sheet, nomeDaGuia) {
  if(nomeDaGuia == "Dashboard"){
    // Será substituído pela função configurarDashboard mais avançada
    sheet.getRange("A1:Z1000").setBackground("#f5f5f5");
    sheet.getRange(1, 1).setValue("Dashboard Interativo - Gestão de Creators");
    sheet.getRange(1, 1).setFontSize(16);
  }
  
  if(nomeDaGuia == "Estatísticas Gerais"){
    // Adiciona fórmulas para cálculos automáticos
    var mainSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Guia Principal");
    
    sheet.getRange("A2").setValue("Total Creators Captados");
    sheet.getRange("B2").setFormula('=COUNTA(\'Guia Principal\'!B2:B)');
    
    sheet.getRange("A3").setValue("Total Reuniões Agendadas");
    sheet.getRange("B3").setFormula('=COUNTIF(\'Guia Principal\'!L2:L, "Reunião Agendada")');
    
    sheet.getRange("A4").setValue("Total Confirmados");
    sheet.getRange("B4").setFormula('=COUNTIF(\'Guia Principal\'!L2:L, "Confirmado")');
    
    sheet.getRange("A5").setValue("Média Engajamento (%)");
    sheet.getRange("B5").setFormula('=AVERAGE(\'Guia Principal\'!G2:G)');
    sheet.getRange("B5").setNumberFormat("0.00%");
    
    sheet.getRange("A6").setValue("Taxa de Conversão (%)");
    sheet.getRange("B6").setFormula('=COUNTIF(\'Guia Principal\'!L2:L, "Parceria Fechada")/COUNTA(\'Guia Principal\'!B2:B)');
    sheet.getRange("B6").setNumberFormat("0.00%");
    
    sheet.getRange("A8").setValue("Creators por Plataforma");
    
    // Tabela de creators por plataforma
    sheet.getRange("A9").setValue("Plataforma");
    sheet.getRange("B9").setValue("Quantidade");
    
    // Preenche plataformas
    var row = 10;
    for (var i = 0; i < PLATAFORMAS.length; i++) {
      sheet.getRange(row, 1).setValue(PLATAFORMAS[i]);
      sheet.getRange(row, 2).setFormula('=COUNTIF(\'Guia Principal\'!E2:E, "' + PLATAFORMAS[i] + '")');
      row++;
    }
    
    // Adiciona cabeçalhos para estatísticas por período
    sheet.getRange("D2").setValue("ANÁLISE POR PERÍODO");
    sheet.getRange("D2").setFontWeight("bold").setBackground("#e3f2fd");
    
    sheet.getRange("D3").setValue("Período");
    sheet.getRange("E3").setValue("Novos Creators");
    sheet.getRange("F3").setValue("Conversões");
    sheet.getRange("G3").setValue("Taxa de Conversão");
    
    // Formata cabeçalhos
    sheet.getRange("D3:G3").setBackground("#f5f5f5").setFontWeight("bold");
    
    // Preenche períodos (últimos 3 meses)
    var today = new Date();
    for (var i = 0; i < 3; i++) {
      var monthDate = new Date(today.getFullYear(), today.getMonth() - i, 1);
      var monthName = Utilities.formatDate(monthDate, "GMT-3", "MMM yyyy");
      var startOfMonth = new Date(monthDate.getFullYear(), monthDate.getMonth(), 1);
      var endOfMonth = new Date(monthDate.getFullYear(), monthDate.getMonth() + 1, 0);
      
      sheet.getRange(4 + i, 4).setValue(monthName);
      
      // Novos creators no mês
      sheet.getRange(4 + i, 5).setFormula('=COUNTIFS(\'Guia Principal\'!K2:K, ">=" & DATE(' + 
                                           monthDate.getFullYear() + ',' + (monthDate.getMonth() + 1) + ',1), \'Guia Principal\'!K2:K, "<=" & DATE(' + 
                                           monthDate.getFullYear() + ',' + (monthDate.getMonth() + 1) + ',' + endOfMonth.getDate() + '))');
      
      // Conversões no mês
      sheet.getRange(4 + i, 6).setFormula('=COUNTIFS(\'Guia Principal\'!L2:L, "Parceria Fechada", \'Guia Principal\'!K2:K, ">=" & DATE(' + 
                                           monthDate.getFullYear() + ',' + (monthDate.getMonth() + 1) + ',1), \'Guia Principal\'!K2:K, "<=" & DATE(' + 
                                           monthDate.getFullYear() + ',' + (monthDate.getMonth() + 1) + ',' + endOfMonth.getDate() + '))');
      
      // Taxa de conversão
      sheet.getRange(4 + i, 7).setFormula('=IF(E' + (4 + i) + '>0, F' + (4 + i) + '/E' + (4 + i) + ', 0)');
      sheet.getRange(4 + i, 7).setNumberFormat("0.00%");
    }
  }
  
  if(nomeDaGuia == "Plano de Conteúdo") {
    // Estrutura para a matriz de conteúdo baseada no ABM Template
    sheet.getRange("A1:K1").merge();
    sheet.getRange("A1").setValue("PLANO DE CONTEÚDO POR JORNADA DO CLIENTE");
    sheet.getRange("A1").setFontSize(14).setFontWeight("bold").setHorizontalAlignment("center");
    sheet.getRange("A1:K1").setBackground("#4285F4").setFontColor("white");
    
    // Cabeçalho da jornada
    var jornadas = [
      "Identificação do Problema",
      "Exploração da Solução",
      "Construção de Requisitos",
      "Seleção",
      "Validação",
      "Criação de Consenso"
    ];
    sheet.getRange(3, 2, 1, jornadas.length).setValues([jornadas]);
    sheet.getRange(3, 2, 1, jornadas.length).setBackground("#e3f2fd").setFontWeight("bold").setHorizontalAlignment("center");
    
    // Tipos de conteúdo
    var tiposConteudo = [
      "Posts Informativos",
      "Tutoriais/How-To",
      "Reviews/Unboxing",
      "Comparativos",
      "Storytelling/Testemunhos",
      "FAQ/Dúvidas Comuns",
      "Behind the Scenes",
      "Lives/Q&A"
    ];
    
    // Inserir tipos de conteúdo
    sheet.getRange(4, 1, tiposConteudo.length, 1).setValues(tiposConteudo.map(tipo => [tipo]));
    sheet.getRange(4, 1, tiposConteudo.length, 1).setBackground("#f5f5f5").setFontWeight("bold");
    
    // Dados da matriz
    var matrizConteudo = [
      // Posts Informativos
      ["• Formato: Posts curtos e objetivos\n• Personalização: Baixa\n• Foco: Awareness inicial"],
      ["• Formato: Posts técnicos\n• Personalização: Média\n• Foco: Educação"],
      ["• Formato: Posts detalhados\n• Personalização: Alta\n• Foco: Especificações"],
      ["• Formato: Posts comparativos\n• Personalização: Alta\n• Foco: Diferenciação"],
      ["• Formato: Posts de validação\n• Personalização: Alta\n• Foco: Confiança"],
      ["• Formato: Posts de consenso\n• Personalização: Alta\n• Foco: Alinhamento"],
      
      // Tutoriais/How-To
      ["• Formato: Vídeos curtos\n• Personalização: Média\n• Foco: Demonstração"],
      ["• Formato: Tutoriais detalhados\n• Personalização: Alta\n• Foco: Aprendizado"],
      ["• Formato: Guias técnicos\n• Personalização: Alta\n• Foco: Implementação"],
      ["• Formato: Tutoriais comparativos\n• Personalização: Alta\n• Foco: Escolha"],
      ["• Formato: Tutoriais avançados\n• Personalização: Alta\n• Foco: Validação"],
      ["• Formato: Tutoriais colaborativos\n• Personalização: Alta\n• Foco: Consenso"],
      
      // Reviews/Unboxing
      ["• Formato: Reviews introdutórios\n• Personalização: Média\n• Foco: Apresentação"],
      ["• Formato: Reviews técnicos\n• Personalização: Alta\n• Foco: Análise"],
      ["• Formato: Reviews detalhados\n• Personalização: Alta\n• Foco: Especificações"],
      ["• Formato: Reviews comparativos\n• Personalização: Alta\n• Foco: Decisão"],
      ["• Formato: Reviews de validação\n• Personalização: Alta\n• Foco: Confirmação"],
      ["• Formato: Reviews de consenso\n• Personalização: Alta\n• Foco: Alinhamento"],
      
      // Comparativos
      ["• Formato: Comparativos básicos\n• Personalização: Média\n• Foco: Introdução"],
      ["• Formato: Comparativos técnicos\n• Personalização: Alta\n• Foco: Análise"],
      ["• Formato: Comparativos detalhados\n• Personalização: Alta\n• Foco: Requisitos"],
      ["• Formato: Comparativos avançados\n• Personalização: Alta\n• Foco: Decisão"],
      ["• Formato: Comparativos de validação\n• Personalização: Alta\n• Foco: Confirmação"],
      ["• Formato: Comparativos de consenso\n• Personalização: Alta\n• Foco: Alinhamento"],
      
      // Storytelling/Testemunhos
      ["• Formato: Histórias introdutórias\n• Personalização: Média\n• Foco: Conexão"],
      ["• Formato: Histórias de solução\n• Personalização: Alta\n• Foco: Exemplos"],
      ["• Formato: Histórias de requisitos\n• Personalização: Alta\n• Foco: Casos"],
      ["• Formato: Histórias de decisão\n• Personalização: Alta\n• Foco: Escolha"],
      ["• Formato: Histórias de validação\n• Personalização: Alta\n• Foco: Confirmação"],
      ["• Formato: Histórias de consenso\n• Personalização: Alta\n• Foco: Alinhamento"],
      
      // FAQ/Dúvidas Comuns
      ["• Formato: FAQ básica\n• Personalização: Baixa\n• Foco: Informação"],
      ["• Formato: FAQ técnica\n• Personalização: Média\n• Foco: Esclarecimento"],
      ["• Formato: FAQ detalhada\n• Personalização: Alta\n• Foco: Requisitos"],
      ["• Formato: FAQ de seleção\n• Personalização: Alta\n• Foco: Decisão"],
      ["• Formato: FAQ de validação\n• Personalização: Alta\n• Foco: Confirmação"],
      ["• Formato: FAQ de consenso\n• Personalização: Alta\n• Foco: Alinhamento"],
      
      // Behind the Scenes
      ["• Formato: BTS básico\n• Personalização: Média\n• Foco: Transparência"],
      ["• Formato: BTS técnico\n• Personalização: Alta\n• Foco: Processo"],
      ["• Formato: BTS detalhado\n• Personalização: Alta\n• Foco: Requisitos"],
      ["• Formato: BTS de seleção\n• Personalização: Alta\n• Foco: Decisão"],
      ["• Formato: BTS de validação\n• Personalização: Alta\n• Foco: Confirmação"],
      ["• Formato: BTS de consenso\n• Personalização: Alta\n• Foco: Alinhamento"],
      
      // Lives/Q&A
      ["• Formato: Lives introdutórias\n• Personalização: Média\n• Foco: Interação"],
      ["• Formato: Lives técnicas\n• Personalização: Alta\n• Foco: Esclarecimento"],
      ["• Formato: Lives de requisitos\n• Personalização: Alta\n• Foco: Detalhamento"],
      ["• Formato: Lives de seleção\n• Personalização: Alta\n• Foco: Decisão"],
      ["• Formato: Lives de validação\n• Personalização: Alta\n• Foco: Confirmação"],
      ["• Formato: Lives de consenso\n• Personalização: Alta\n• Foco: Alinhamento"]
    ];
    
    // Inserir dados da matriz
    sheet.getRange(4, 2, matrizConteudo.length, matrizConteudo[0].length).setValues(matrizConteudo);
    
    // Formatação condicional para níveis de personalização
    var niveis = ["Baixa", "Média", "Alta"];
    var cores = ["#c8e6c9", "#fff9c4", "#ffcdd2"];
    
    for (var i = 0; i < niveis.length; i++) {
      var range = sheet.getRange(4, 2, matrizConteudo.length, matrizConteudo[0].length);
      var rule = SpreadsheetApp.newConditionalFormatRule()
        .whenTextContains(niveis[i])
        .setBackground(cores[i])
        .build();
      var rules = [rule];
      range.setConditionalFormatRules(rules);
    }
    
    // Ajustar altura das linhas
    sheet.setRowHeights(4, matrizConteudo.length, 100);
    
    // Adicionar instruções
    var lastRow = 4 + matrizConteudo.length + 2;
    sheet.getRange(lastRow, 1, 1, jornadas.length + 1).merge();
    sheet.getRange(lastRow, 1).setValue("Instruções: Use esta matriz para planejar o conteúdo ideal para cada fase da jornada do cliente. Considere o formato e nível de personalização necessários para cada interseção.");
    sheet.getRange(lastRow, 1).setFontStyle("italic").setBackground("#fff3e0");
    
    // Congelar cabeçalhos
    sheet.setFrozenRows(3);
    sheet.setFrozenColumns(1);
  }
  
  if(nomeDaGuia == "Qualificação dos Creators") {
    // Configuração básica da planilha
    sheet.setColumnWidth(1, 200);  // Nome Creator
    sheet.setColumnWidth(2, 150);  // Potencial Alcance
    sheet.setColumnWidth(3, 150);  // Relevância para Audiência
    sheet.setColumnWidth(4, 150);  // Compatibilidade com Marca
    sheet.setColumnWidth(5, 150);  // Qualidade de Conteúdo
    sheet.setColumnWidth(6, 150);  // Engajamento da Audiência
    sheet.setColumnWidth(7, 150);  // Histórico de Conversão
    sheet.setColumnWidth(8, 150);  // Pontuação Total
    sheet.setColumnWidth(9, 150);  // Status Qualificação
    sheet.setColumnWidth(10, 150); // Segmento
    sheet.setColumnWidth(11, 150); // Tier de Prioridade
    
    // Novos campos para Capacidade de Personalização
    sheet.setColumnWidth(12, 150); // Flexibilidade Adaptação
    sheet.setColumnWidth(13, 150); // Autenticidade Conteúdo
    sheet.setColumnWidth(14, 150); // Consistência Briefings
    sheet.setColumnWidth(15, 150); // Conhecimento Produto
    sheet.setColumnWidth(16, 150); // Adequação Segmentos
    sheet.setColumnWidth(17, 150); // Pontuação Personalização
    
    // Título principal
    sheet.getRange("A1:Q1").merge();
    sheet.getRange("A1").setValue("QUALIFICAÇÃO DOS CREATORS");
    sheet.getRange("A1").setFontSize(16).setFontWeight("bold").setHorizontalAlignment("center");
    sheet.getRange("A1:Q1").setBackground("#4285F4").setFontColor("white");
    
    // Subtítulo
    sheet.getRange("A2:Q2").merge();
    sheet.getRange("A2").setValue("Avaliação Completa de Potencial e Capacidade de Personalização");
    sheet.getRange("A2").setHorizontalAlignment("center").setFontStyle("italic");
    
    // Cabeçalhos das colunas
    var headers = [
      "Nome Creator",
      "Potencial Alcance (1-5)",
      "Relevância para Audiência (1-5)",
      "Compatibilidade com Marca (1-5)",
      "Qualidade de Conteúdo (1-5)",
      "Engajamento da Audiência (1-5)",
      "Histórico de Conversão (1-5)",
      "Pontuação Total",
      "Status Qualificação",
      "Segmento",
      "Tier de Prioridade",
      "Flexibilidade Adaptação (1-5)",
      "Autenticidade Conteúdo (1-5)",
      "Consistência Briefings (1-5)",
      "Conhecimento Produto (1-5)",
      "Adequação Segmentos (1-5)",
      "Pontuação Personalização"
    ];
    
    sheet.getRange(4, 1, 1, headers.length).setValues([headers]);
    sheet.getRange(4, 1, 1, headers.length).setBackground("#e3f2fd").setFontWeight("bold");
    
    // Validação para pontuações (1-5)
    var rangeValidation = SpreadsheetApp.newDataValidation()
      .requireNumberBetween(1, 5)
      .setAllowInvalid(false)
      .build();
    sheet.getRange("B5:G1000").setDataValidation(rangeValidation);
    sheet.getRange("L5:P1000").setDataValidation(rangeValidation);
    
    // Fórmula para pontuação total
    sheet.getRange("H5").setFormula("=SUM(B5:G5)");
    sheet.getRange("H5:H1000").setFormula("=SUM(B5:G5)");
    
    // Fórmula para pontuação de personalização
    sheet.getRange("Q5").setFormula("=SUM(L5:P5)");
    sheet.getRange("Q5:Q1000").setFormula("=SUM(L5:P5)");
    
    // Formatação condicional para pontuações
    var pontuacaoRange = sheet.getRange("H5:H1000");
    var rule1 = SpreadsheetApp.newConditionalFormatRule()
      .whenNumberGreaterThan(20)
      .setBackground("#c8e6c9")
      .build();
    var rule2 = SpreadsheetApp.newConditionalFormatRule()
      .whenNumberBetween(15, 20)
      .setBackground("#fff9c4")
      .build();
    var rule3 = SpreadsheetApp.newConditionalFormatRule()
      .whenNumberLessThan(15)
      .setBackground("#ffcdd2")
      .build();
    pontuacaoRange.setConditionalFormatRules([rule1, rule2, rule3]);
    
    // Formatação condicional para pontuação de personalização
    var personalizacaoRange = sheet.getRange("Q5:Q1000");
    var rule4 = SpreadsheetApp.newConditionalFormatRule()
      .whenNumberGreaterThan(20)
      .setBackground("#c8e6c9")
      .build();
    var rule5 = SpreadsheetApp.newConditionalFormatRule()
      .whenNumberBetween(15, 20)
      .setBackground("#fff9c4")
      .build();
    var rule6 = SpreadsheetApp.newConditionalFormatRule()
      .whenNumberLessThan(15)
      .setBackground("#ffcdd2")
      .build();
    personalizacaoRange.setConditionalFormatRules([rule4, rule5, rule6]);
    
    // Adicionar instruções
    var lastRow = 1000;
    sheet.getRange(lastRow, 1, 1, headers.length).merge();
    sheet.getRange(lastRow, 1).setValue("Instruções: Avalie cada creator em uma escala de 1 a 5 para todos os critérios. A pontuação total e de personalização ajudarão a identificar os creators mais adequados para diferentes tipos de campanhas.");
    sheet.getRange(lastRow, 1).setFontStyle("italic").setBackground("#fff3e0");
    
    // Congelar cabeçalhos
    sheet.setFrozenRows(4);
    sheet.setFrozenColumns(1);
  }
}

/**
 * Inicializa a guia de Análise de Interesses com caminhos de engajamento por segmento
 */
function inicializarAnaliseInteresses() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("Análise de Interesses");
  
  if (!sheet) {
    return;
  }
  
  sheet.clear();
  
  // Configuração básica da planilha
  sheet.setColumnWidth(1, 30);  // Segmento
  sheet.setColumnWidth(2, 200); // Creators Recomendados
  sheet.setColumnWidth(3, 200); // Sequência de Conteúdo
  sheet.setColumnWidth(4, 200); // Pontos de Contato
  sheet.setColumnWidth(5, 200); // Gatilhos de Avanço
  sheet.setColumnWidth(6, 200); // Métricas de Sucesso
  
  // Título principal
  sheet.getRange("A1:F1").merge();
  sheet.getRange("A1").setValue("CAMINHOS DE ENGAJAMENTO POR SEGMENTO");
  sheet.getRange("A1").setFontSize(16).setFontWeight("bold").setHorizontalAlignment("center");
  sheet.getRange("A1:F1").setBackground("#4285F4").setFontColor("white");
  
  // Subtítulo
  sheet.getRange("A2:F2").merge();
  sheet.getRange("A2").setValue("Mapeamento de Jornadas Personalizadas por Perfil de Audiência");
  sheet.getRange("A2").setHorizontalAlignment("center").setFontStyle("italic");
  
  // Cabeçalhos das colunas
  var headers = [
    "Segmento",
    "Creators Recomendados",
    "Sequência de Conteúdo",
    "Pontos de Contato",
    "Gatilhos de Avanço",
    "Métricas de Sucesso"
  ];
  
  sheet.getRange(4, 1, 1, headers.length).setValues([headers]);
  sheet.getRange(4, 1, 1, headers.length).setBackground("#e3f2fd").setFontWeight("bold");
  
  // Dados dos segmentos
  var segmentos = [
    // Iniciantes
    [
      "Iniciantes",
      "• Creators educacionais\n• Micro-influencers\n• Tutores especializados\n• Mentores da área",
      "1. Conteúdo introdutório\n2. Tutoriais básicos\n3. Dicas práticas\n4. Casos de sucesso simples\n5. Guias passo a passo",
      "• Instagram Stories\n• Posts informativos\n• Lives semanais\n• Grupos de WhatsApp\n• Newsletters básicas",
      "• Primeira interação com conteúdo\n• Compartilhamento de dúvidas\n• Participação em lives\n• Salvamento de posts\n• Inscrição em newsletter",
      "• Taxa de engajamento inicial\n• Tempo médio de visualização\n• Taxa de salvamento\n• Número de dúvidas\n• Taxa de inscrição"
    ],
    
    // Entusiastas
    [
      "Entusiastas",
      "• Creators intermediários\n• Especialistas da área\n• Comunidades ativas\n• Referências do setor",
      "1. Conteúdo avançado\n2. Análises técnicas\n3. Tendências do mercado\n4. Casos de sucesso complexos\n5. Workshops especializados",
      "• Lives técnicas\n• Webinars\n• Grupos VIP\n• Eventos presenciais\n• Mentoria em grupo",
      "• Participação em workshops\n• Compartilhamento de experiências\n• Criação de conteúdo\n• Interação em grupos VIP\n• Participação em eventos",
      "• Taxa de participação em eventos\n• Qualidade das interações\n• Número de compartilhamentos\n• Engajamento em grupos\n• Taxa de conversão para VIP"
    ],
    
    // Profissionais
    [
      "Profissionais",
      "• Creators premium\n• Líderes de mercado\n• Especialistas renomados\n• Influenciadores estratégicos",
      "1. Conteúdo estratégico\n2. Análises de mercado\n3. Tendências globais\n4. Casos de sucesso empresariais\n5. Consultorias especializadas",
      "• Masterminds\n• Consultorias individuais\n• Eventos exclusivos\n• Redes de networking\n• Parcerias estratégicas",
      "• Participação em masterminds\n• Implementação de estratégias\n• Geração de resultados\n• Networking ativo\n• Parcerias comerciais",
      "• ROI das estratégias\n• Taxa de implementação\n• Resultados gerados\n• Expansão de rede\n• Valor das parcerias"
    ]
  ];
  
  // Inserir dados dos segmentos
  sheet.getRange(5, 1, segmentos.length, segmentos[0].length).setValues(segmentos);
  
  // Formatação condicional para segmentos
  var cores = ["#e3f2fd", "#f3e5f5", "#e8f5e9"];
  for (var i = 0; i < segmentos.length; i++) {
    sheet.getRange(5 + i, 1, 1, segmentos[0].length).setBackground(cores[i]);
  }
  
  // Adicionar instruções
  var lastRow = 5 + segmentos.length + 2;
  sheet.getRange(lastRow, 1, 1, headers.length).merge();
  sheet.getRange(lastRow, 1).setValue("Instruções: Use esta matriz para planejar e acompanhar os caminhos de engajamento específicos para cada segmento. Atualize as métricas regularmente para otimizar as estratégias.");
  sheet.getRange(lastRow, 1).setFontStyle("italic").setBackground("#fff3e0");
  
  // Congelar cabeçalhos
  sheet.setFrozenRows(4);
  sheet.setFrozenColumns(1);
}

/**
 * Inicializa a guia Playbook de Engajamento com táticas e estratégias
 */
function inicializarPlaybookEngajamento() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("Playbook de Engajamento");
  
  if (!sheet) {
    return;
  }
  
  sheet.clear();
  
  // Configuração básica da planilha
  sheet.setColumnWidth(1, 200);  // Tática
  sheet.setColumnWidth(2, 150);  // Fase
  sheet.setColumnWidth(3, 200);  // Creator Ideal
  sheet.setColumnWidth(4, 150);  // Personalização
  sheet.setColumnWidth(5, 200);  // Objetivo
  sheet.setColumnWidth(6, 150);  // Métrica
  sheet.setColumnWidth(7, 150);  // Orçamento
  
  // Título principal
  sheet.getRange("A1:G1").merge();
  sheet.getRange("A1").setValue("PLAYBOOK DE ENGAJAMENTO");
  sheet.getRange("A1").setFontSize(16).setFontWeight("bold").setHorizontalAlignment("center");
  sheet.getRange("A1:G1").setBackground("#4285F4").setFontColor("white");
  
  // Subtítulo
  sheet.getRange("A2:G2").merge();
  sheet.getRange("A2").setValue("Estratégias e Táticas para Cada Fase da Jornada do Cliente");
  sheet.getRange("A2").setHorizontalAlignment("center").setFontStyle("italic");
  
  // Cabeçalhos das colunas
  var headers = [
    "Tática",
    "Fase",
    "Creator Ideal",
    "Personalização",
    "Objetivo",
    "Métrica",
    "Orçamento"
  ];
  
  sheet.getRange(4, 1, 1, headers.length).setValues([headers]);
  sheet.getRange(4, 1, 1, headers.length).setBackground("#e3f2fd").setFontWeight("bold");
  
  // Dados das táticas
  var taticas = [
    // Identificação do Problema
    [
      "Anúncios de display para exploração inicial do problema",
      "Identificação do Problema",
      "• Creators especializados em educação\n• Micro-influencers do setor\n• Especialistas em conteúdo técnico",
      "Baixa",
      "Aumentar awareness sobre desafios comuns do setor",
      "• CTR dos anúncios\n• Taxa de engajamento\n• Tempo médio de visualização",
      "R$ 5.000 - R$ 10.000"
    ],
    [
      "Direct mail com casos de uso semelhantes",
      "Identificação do Problema",
      "• Creators com experiência em storytelling\n• Especialistas em casos de sucesso\n• Influenciadores do setor",
      "Média",
      "Demonstrar relevância através de exemplos práticos",
      "• Taxa de abertura\n• Taxa de resposta\n• Número de agendamentos",
      "R$ 3.000 - R$ 7.000"
    ],
    
    // Exploração da Solução
    [
      "Posts em redes sociais com benchmarks do setor",
      "Exploração da Solução",
      "• Creators de dados e analytics\n• Especialistas em mercado\n• Influenciadores de tendências",
      "Média",
      "Educar sobre padrões e melhores práticas do setor",
      "• Engajamento por post\n• Taxa de compartilhamento\n• Comentários qualificados",
      "R$ 2.000 - R$ 5.000"
    ],
    [
      "Webinars para demonstração de solução",
      "Exploração da Solução",
      "• Especialistas técnicos\n• Creators de conteúdo educacional\n• Influenciadores do setor",
      "Alta",
      "Apresentar soluções de forma interativa",
      "• Número de inscritos\n• Taxa de participação\n• NPS do webinar",
      "R$ 8.000 - R$ 15.000"
    ],
    
    // Construção de Requisitos
    [
      "Eventos de networking para decisores",
      "Construção de Requisitos",
      "• Creators de networking\n• Líderes de mercado\n• Especialistas em relacionamento",
      "Alta",
      "Facilitar conexões e alinhamento de requisitos",
      "• Número de participantes\n• Qualidade das interações\n• Leads gerados",
      "R$ 10.000 - R$ 20.000"
    ],
    [
      "Conteúdo segmentado para diferentes stakeholders",
      "Construção de Requisitos",
      "• Creators especializados em B2B\n• Especialistas em comunicação\n• Influenciadores setoriais",
      "Alta",
      "Endereçar necessidades específicas de cada stakeholder",
      "• Engajamento por segmento\n• Taxa de conversão\n• Feedback qualitativo",
      "R$ 5.000 - R$ 12.000"
    ],
    
    // Seleção
    [
      "Mensagens personalizadas em LinkedIn",
      "Seleção",
      "• Creators de LinkedIn\n• Especialistas em vendas\n• Influenciadores B2B",
      "Alta",
      "Gerar conversas qualificadas com decisores",
      "• Taxa de resposta\n• Taxa de agendamento\n• Qualidade dos leads",
      "R$ 3.000 - R$ 8.000"
    ]
  ];
  
  // Inserir dados das táticas
  sheet.getRange(5, 1, taticas.length, taticas[0].length).setValues(taticas);
  
  // Formatação condicional para níveis de personalização
  var personalizacaoRange = sheet.getRange("D5:D" + (5 + taticas.length - 1));
  var rule1 = SpreadsheetApp.newConditionalFormatRule()
    .whenTextEqualTo("Baixa")
    .setBackground("#c8e6c9")
    .build();
  var rule2 = SpreadsheetApp.newConditionalFormatRule()
    .whenTextEqualTo("Média")
    .setBackground("#fff9c4")
    .build();
  var rule3 = SpreadsheetApp.newConditionalFormatRule()
    .whenTextEqualTo("Alta")
    .setBackground("#ffcdd2")
    .build();
  personalizacaoRange.setConditionalFormatRules([rule1, rule2, rule3]);
  
  // Formatação condicional para fases
  var faseRange = sheet.getRange("B5:B" + (5 + taticas.length - 1));
  var cores = ["#e3f2fd", "#f3e5f5", "#e8f5e9", "#fff3e0"];
  var fases = ["Identificação do Problema", "Exploração da Solução", "Construção de Requisitos", "Seleção"];
  
  for (var i = 0; i < fases.length; i++) {
    var rule = SpreadsheetApp.newConditionalFormatRule()
      .whenTextEqualTo(fases[i])
      .setBackground(cores[i])
      .build();
    faseRange.setConditionalFormatRules([rule]);
  }
  
  // Adicionar instruções
  var lastRow = 5 + taticas.length + 2;
  sheet.getRange(lastRow, 1, 1, headers.length).merge();
  sheet.getRange(lastRow, 1).setValue("Instruções: Use este playbook para planejar e executar estratégias de engajamento em cada fase da jornada. Considere o nível de personalização necessário e selecione os creators mais adequados para cada tática.");
  sheet.getRange(lastRow, 1).setFontStyle("italic").setBackground("#fff3e0");
  
  // Congelar cabeçalhos
  sheet.setFrozenRows(4);
  sheet.setFrozenColumns(1);
}
