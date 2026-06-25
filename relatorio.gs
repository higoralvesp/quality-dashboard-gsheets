// ============================================================
//  RELATÓRIO DE QUALIDADE — BACKOFFICE v2
//  - Relatório mensal completo (HTML 1)
//  - Diagnóstico por período selecionado (HTML 2)
//  - Texto para grupo (e-mail separado)
// ============================================================

var EMAIL_SUPERVISAO = "supervisao@empresa.com"; // e-mail da supervisão
var EMAIL_GRUPO      = "seu@email.com";          // seu e-mail para texto do grupo
var NOME_ABA_PF      = "PENTE FINO";
var NOME_ABA_DV      = "DADOS VALIDAÇÕES";

var MESES_PT = ["Janeiro","Fevereiro","Março","Abril","Maio","Junho",
                "Julho","Agosto","Setembro","Outubro","Novembro","Dezembro"];

function mesAnoAtual() {
  var d = new Date();
  return MESES_PT[d.getMonth()] + " " + d.getFullYear();
}

function formatarData(d) {
  return Utilities.formatDate(d, "America/Sao_Paulo", "dd/MM/yyyy");
}

// ── TIPOS VÁLIDOS (minúsculo para comparação case-insensitive) ──
var TIPOS_COL_N_LOWER = [
  "adesão", "adesão - fmv", "adesão/reativação",
  "adesão/reativação - p.f", "adesão/reativação c,e",
  "adesão - consultor externo", "substituição/reativação"
];
var TIPO_REATIVACAO_LOWER     = "reativação";
var TIPO_REATIVACAO_CRM_LOWER = "reativação crm";
var TIPOS_COL_W_LOWER = ["substituição", "substituição/reativação"];

// ── ÍNDICES DAS COLUNAS (base 0) ──
var C_TIPO  = 12; // M — tipo de boleto
var C_USU   = 13; // N — usuário da baixa
var C_CLASS = 20; // U — classificação
var C_AGT_W = 22; // W — agente substituição
var C_DATA  = 21; // V — data da análise
var C_ERRO  = 26; // AA — tipo do erro
var C_GRAV  = 29; // AD — gravidade
var C_FASE  = 30; // AE — fase
var C_RESP  = 31; // AF — agente responsável pelo erro

// ============================================================
//  FUNÇÃO PRINCIPAL
// ============================================================
function gerarRelatorio() {
  var ui = SpreadsheetApp.getUi();
  var ss = SpreadsheetApp.getActiveSpreadsheet();

  // ── PERÍODO DO DIAGNÓSTICO (automático) ─────────────────────
  // Segunda (1):      semana anterior — seg a sab (7 dias atrás até 1 dia atrás)
  // Terça a sexta (2-5): semana atual — seg até ontem
  // Sábado (6):       semana atual — seg a sex (semana completa)
  var dataInicioFiltro, dataFimFiltro;
  var hoje      = new Date();
  var diaSemana = hoje.getDay(); // 0=dom, 1=seg, 2=ter... 6=sab

  if (diaSemana === 1) {
    // Segunda: pega semana anterior (seg a sab)
    dataInicioFiltro = new Date(hoje);
    dataInicioFiltro.setDate(hoje.getDate() - 7); // segunda anterior
    dataInicioFiltro.setHours(0,0,0,0);
    dataFimFiltro = new Date(hoje);
    dataFimFiltro.setDate(hoje.getDate() - 1); // sábado anterior
    dataFimFiltro.setHours(23,59,59,999);

  } else if (diaSemana === 6) {
    // Sábado: pega semana atual completa (seg a sex)
    dataInicioFiltro = new Date(hoje);
    dataInicioFiltro.setDate(hoje.getDate() - 5); // segunda atual
    dataInicioFiltro.setHours(0,0,0,0);
    dataFimFiltro = new Date(hoje);
    dataFimFiltro.setDate(hoje.getDate() - 1); // sexta atual
    dataFimFiltro.setHours(23,59,59,999);

  } else {
    // Terça a sexta: semana atual, segunda até ontem
    var diasDesdeSegunda = diaSemana - 1;
    dataInicioFiltro = new Date(hoje);
    dataInicioFiltro.setDate(hoje.getDate() - diasDesdeSegunda); // segunda atual
    dataInicioFiltro.setHours(0,0,0,0);
    dataFimFiltro = new Date(hoje);
    dataFimFiltro.setDate(hoje.getDate() - 1); // ontem
    dataFimFiltro.setHours(23,59,59,999);
  }

  // ── LER PLANILHA ────────────────────────────────────────────
  var wsPF = ss.getSheetByName(NOME_ABA_PF);
  var wsDV = ss.getSheetByName(NOME_ABA_DV);
  if (!wsPF) { ui.alert('Aba "' + NOME_ABA_PF + '" não encontrada.'); return; }
  if (!wsDV) { ui.alert('Aba "' + NOME_ABA_DV + '" não encontrada.'); return; }

  SpreadsheetApp.flush();
  var dados   = wsPF.getDataRange().getValues();
  var dvDados = wsDV.getDataRange().getValues();
  var hoje    = formatarData(new Date());

  // ── MAPA DE NOMES ────────────────────────────────────────────
  var nomeParaBKO = {};
  for (var i = 1; i < dvDados.length; i++) {
    var nc  = dvDados[i][4] ? String(dvDados[i][4]).trim().toUpperCase() : "";
    var nbk = dvDados[i][5] ? String(dvDados[i][5]).trim() : "";
    if (nc && nbk) nomeParaBKO[nc] = nbk;
  }

  // ── PROCESSAR DADOS COMPLETOS (mês inteiro) ──────────────────
  var dadosMes     = processarDados(dados, nomeParaBKO, null, null);

  // ── PROCESSAR DADOS DO PERÍODO (diagnóstico) ────────────────
  var dadosPeriodo = processarDados(dados, nomeParaBKO, dataInicioFiltro, dataFimFiltro);

  if (dadosMes.totalErros === 0) {
    ui.alert("Nenhum erro encontrado. Verifique se os dados estão atualizados.");
    return;
  }

  // ── PERÍODO DO MÊS (para o relatório mensal) ─────────────────
  var periodoMes = (dadosMes.dataInicio && dadosMes.dataFim)
    ? dadosMes.dataInicio + ' a ' + dadosMes.dataFim
    : mesAnoAtual();

  // ── PERÍODO DO DIAGNÓSTICO ───────────────────────────────────
  var periodoDiag = formatarData(dataInicioFiltro) + ' a ' + formatarData(dataFimFiltro);

  // ── GERAR OS DOIS HTMLs ──────────────────────────────────────
  var htmlMes     = gerarHTML(hoje, dadosMes, periodoMes, "Consolidado Mensal");
  var htmlDiag    = gerarHTML(hoje, dadosPeriodo, periodoDiag, "Diagnóstico do Período");

  // ── GERAR TEXTO PARA O GRUPO ─────────────────────────────────
  var textoGrupo  = gerarTextoGrupo(dadosPeriodo, periodoDiag);

  // ── NOME DOS ARQUIVOS ────────────────────────────────────────
  var mesRef  = Utilities.formatDate(new Date(), "America/Sao_Paulo", "MM_yyyy");
  var diagRef = formatarData(dataInicioFiltro).replace(/\//g,'-') + '_a_' + formatarData(dataFimFiltro).replace(/\//g,'-');

  // ── RASCUNHO: SUPERVISÃO (dois HTMLs em anexo) ──────────────
  GmailApp.createDraft(
    EMAIL_SUPERVISAO,
    "Relatório de Qualidade — Backoffice | " + mesAnoAtual(),
    "",
    {
      htmlBody: '<p style="font-family:Arial,sans-serif;font-size:14px;color:#333;">Bom dia, seguem em anexo o relatório mensal e o diagnóstico do período de <strong>' + periodoDiag + '</strong>.</p>',
      attachments: [
        Utilities.newBlob(htmlMes,  "text/html", "relatorio_bko_"  + mesRef  + ".html"),
        Utilities.newBlob(htmlDiag, "text/html", "diagnostico_bko_" + diagRef + ".html")
      ]
    }
  );

  // ── E-MAIL: VOCÊ (texto para o grupo — enviado direto) ───────
  GmailApp.sendEmail(
    EMAIL_GRUPO,
    "Qualidade | Texto para o grupo — " + periodoDiag,
    textoGrupo
  );

  ui.alert("✅ Pronto!\n\nRascunho criado no Gmail — revise e envie para a supervisão.\nTexto para o grupo enviado para " + EMAIL_GRUPO + ".");
}

// ============================================================
//  PROCESSAR DADOS (com ou sem filtro de data)
// ============================================================
function processarDados(dados, nomeParaBKO, dataInicio, dataFim) {
  var analisadosPorAgente = {};
  var erros = [];
  var datas = [];

  for (var i = 1; i < dados.length; i++) {
    var row    = dados[i];
    var tipo   = row[C_TIPO]  ? String(row[C_TIPO]).trim()           : "";
    var tipoLow = tipo.toLowerCase();
    var usuN   = row[C_USU]   ? String(row[C_USU]).trim().toUpperCase()  : "";
    var classU = row[C_CLASS] ? String(row[C_CLASS]).trim().toLowerCase(): "";
    var usuW   = row[C_AGT_W] ? String(row[C_AGT_W]).trim().toUpperCase(): "";
    var dataV  = row[C_DATA];
    var resp   = row[C_RESP]  ? String(row[C_RESP]).trim()           : "";

    // Filtro de data (se informado)
    if (dataInicio && dataFim) {
      if (!(dataV instanceof Date) || dataV < dataInicio || dataV > dataFim) continue;
    }

    // Coletar datas para período
    if (dataV instanceof Date && !isNaN(dataV.getTime())) datas.push(dataV);

    // Contar analisados
    var agenteBKO = "";
    if (usuN && TIPOS_COL_N_LOWER.indexOf(tipoLow) !== -1) {
      agenteBKO = nomeParaBKO[usuN] || "";
    }
    if (!agenteBKO && usuN && tipoLow === TIPO_REATIVACAO_LOWER && classU === TIPO_REATIVACAO_CRM_LOWER) {
      agenteBKO = nomeParaBKO[usuN] || "";
    }
    if (!agenteBKO && usuW && TIPOS_COL_W_LOWER.indexOf(tipoLow) !== -1) {
      agenteBKO = nomeParaBKO[usuW] || "";
    }
    if (agenteBKO) {
      analisadosPorAgente[agenteBKO] = (analisadosPorAgente[agenteBKO] || 0) + 1;
    }

    // Coletar erros (filtrar por AF)
    if (resp !== "") {
      var erro = row[C_ERRO] ? String(row[C_ERRO]).trim() : "";
      var grav = row[C_GRAV] ? String(row[C_GRAV]).trim() : "-";
      var fase = row[C_FASE] ? String(row[C_FASE]).trim() : "-";
      erros.push({ tipo: erro, gravidade: grav, fase: fase, agente: resp });
    }
  }

  // Consolidado geral por tipo
  var tipoMap = {};
  erros.forEach(function(e) {
    if (!tipoMap[e.tipo]) tipoMap[e.tipo] = { gravidade: e.gravidade, fase: e.fase, qtd: 0 };
    tipoMap[e.tipo].qtd++;
  });
  var tipoArr = Object.keys(tipoMap).map(function(t) {
    return { tipo: t, gravidade: tipoMap[t].gravidade, fase: tipoMap[t].fase, qtd: tipoMap[t].qtd };
  }).sort(function(a,b){ return b.qtd - a.qtd; });

  // Consolidado por agente
  var agenteErros = {};
  erros.forEach(function(e) {
    if (!e.agente) return;
    if (!agenteErros[e.agente]) agenteErros[e.agente] = { erros: {}, total: 0 };
    agenteErros[e.agente].erros[e.tipo] = (agenteErros[e.agente].erros[e.tipo] || 0) + 1;
    agenteErros[e.agente].total++;
  });
  var agentesArr = Object.keys(agenteErros).map(function(ag) {
    var analisados = analisadosPorAgente[ag] || 0;
    var totalErros = agenteErros[ag].total;
    var taxa = analisados > 0 ? (totalErros / analisados) * 100 : 0;
    return { nome: ag, analisados: analisados, totalErros: totalErros, taxa: taxa, erros: agenteErros[ag].erros };
  }).sort(function(a,b){ return b.taxa - a.taxa; });

  // Métricas
  var totalAnalisado = Object.keys(analisadosPorAgente).reduce(function(s,k){ return s+analisadosPorAgente[k]; }, 0);
  var totalErros     = erros.length;
  var taxaGeral      = totalAnalisado > 0 ? ((totalErros/totalAnalisado)*100).toFixed(1) : "0.0";
  var gravissimos    = erros.filter(function(e){ return e.gravidade.toUpperCase()==="GRAVÍSSIMO"; }).length;

  datas.sort(function(a,b){ return a-b; });
  var dataInicioStr = datas.length > 0 ? formatarData(datas[0])            : "";
  var dataFimStr    = datas.length > 0 ? formatarData(datas[datas.length-1]): "";

  return {
    totalAnalisado: totalAnalisado,
    totalErros:     totalErros,
    taxaGeral:      taxaGeral,
    gravissimos:    gravissimos,
    nAgentes:       agentesArr.length,
    tipoArr:        tipoArr,
    tipoMap:        tipoMap,
    agentesArr:     agentesArr,
    dataInicio:     dataInicioStr,
    dataFim:        dataFimStr
  };
}

// ============================================================
//  TEXTO PARA O GRUPO
// ============================================================
function gerarTextoGrupo(d, periodo) {
  var linhas = [];
  linhas.push("📊 QUALIDADE BACKOFFICE — " + mesAnoAtual());
  linhas.push("Período: " + periodo);
  linhas.push(d.totalAnalisado + " análises realizadas | " + d.totalErros + " erros | Taxa geral: " + d.taxaGeral + "%");
  linhas.push("");

  // Resumo geral
  var top1 = d.tipoArr.length > 0 ? d.tipoArr[0] : null;
  var top2 = d.tipoArr.length > 1 ? d.tipoArr[1] : null;

  var resumo = "Encerramos o período com " + d.totalErros + " erros identificados em " + d.totalAnalisado + " análises realizadas, representando uma taxa geral de " + d.taxaGeral + "%.";
  if (d.gravissimos > 0) {
    resumo += " Do total, " + d.gravissimos + " foram classificados como gravíssimos, o que exige atenção redobrada da equipe.";
  }
  linhas.push(resumo);
  linhas.push("");

  if (top1) {
    var erroTexto = "O erro mais recorrente do período foi " + top1.tipo + " com " + top1.qtd + " ocorrência(s)";
    if (top2) erroTexto += ", seguido de " + top2.tipo + " com " + top2.qtd + " caso(s)";
    erroTexto += ". Ambos indicam falha no preenchimento completo do cadastro antes da exportação.";
    linhas.push(erroTexto);
    linhas.push("");
  }

  // Resultado por agente
  linhas.push("Resultado por agente:");
  d.agentesArr.forEach(function(ag) {
    var nome = ag.nome.replace("BKO - ", "");
    var emoji = ag.taxa >= 35 ? "🔴" : ag.taxa >= 20 ? "🟠" : "🟢";
    linhas.push(emoji + " " + nome + " → " + ag.totalErros + " erros em " + ag.analisados + " análises (" + ag.taxa.toFixed(1) + "%)");
  });
  linhas.push("");

  // Ponto focal
  var principaisErros = d.tipoArr.slice(0,3).map(function(t){ return t.tipo; }).join(", ");
  linhas.push("Ponto focal da equipe para o próximo período: atenção ao preenchimento completo das informações no momento da ativação. Os principais erros foram: " + principaisErros + ".");

  return linhas.join("\n");
}

// ============================================================
//  HELPERS
// ============================================================
function gerarNota(agente) {
  var taxa = agente.taxa.toFixed(1);
  var topErros = Object.keys(agente.erros).map(function(t){
    return { tipo: t, qtd: agente.erros[t] };
  }).sort(function(a,b){ return b.qtd - a.qtd; });

  if (topErros.length === 0) return "Nenhum erro registrado no período.";
  var principal = topErros[0];
  var nota = "";

  if (agente.taxa >= 40)      nota += "Taxa de " + taxa + "% exige atenção imediata. ";
  else if (agente.taxa >= 20) nota += "Taxa de " + taxa + "% está acima da meta e merece atenção. ";
  else                        nota += "Taxa de " + taxa + "% está em nível controlado, mas há pontos a melhorar. ";

  if (agente.analisados > 0 && agente.analisados < 50) {
    nota += "Amostra de apenas " + agente.analisados + " registros pode amplificar o percentual. Ainda assim, ";
  }

  var t = principal.tipo.toLowerCase();
  if (t.indexOf("inf. secund") !== -1 || t.indexOf("secundári") !== -1) {
    nota += "o padrão é concentrado no <strong>preenchimento de informações secundárias</strong> — chassi, classificação e dados complementares. Ponto focal: atenção ao cadastro completo no momento da ativação, campo a campo.";
  } else if (t.indexOf("chassi") !== -1 || t.indexOf("fipe") !== -1) {
    nota += "o principal erro é <strong>Chassi/Modelo FIPE</strong> com " + principal.qtd + " ocorrência(s). Ponto focal: conferir os dados do veículo na ativação antes de exportar.";
  } else if (t.indexOf("e-mail") !== -1 || t.indexOf("email") !== -1) {
    nota += "destaque para <strong>e-mail inválido ou não preenchido</strong> com " + principal.qtd + " caso(s). Ponto focal: validar o e-mail do associado antes de finalizar o cadastro.";
  } else if (t.indexOf("classifica") !== -1) {
    nota += "o erro mais recorrente é <strong>Classificação/ID em branco ou data contrato</strong>. Ponto focal: revisar os campos de identificação na exportação.";
  } else if (t.indexOf("termo") !== -1) {
    nota += "o <strong>Termo não enviado</strong> aparece como principal ocorrência. Ponto focal: verificar o envio do termo ao associado antes de concluir o processo.";
  } else if (t.indexOf("taxa adm") !== -1 || t.indexOf("ajuste") !== -1) {
    nota += "destaque para <strong>Ajuste de taxa administrativa não lançada</strong>. Ponto focal: conferir o lançamento da taxa antes da exportação do boleto.";
  } else if (t.indexOf("assist") !== -1 || t.indexOf("residencial") !== -1) {
    nota += "o erro <strong>Assistência Residencial Improcedente</strong> lidera as ocorrências. Ponto focal: verificar a elegibilidade do associado antes de acionar o benefício.";
  } else if (t.indexOf("endere") !== -1) {
    nota += "destaque para <strong>Endereço do Veículo</strong> incorreto ou em branco. Ponto focal: conferir o endereço registrado no momento da ativação.";
  } else if (t.indexOf("boleto") !== -1) {
    nota += "o principal erro envolve <strong>Boleto</strong> (" + principal.tipo + "). Ponto focal: revisar geração, valores e transferência dos boletos.";
  } else {
    nota += principal.qtd + " ocorrência(s) de <strong>" + principal.tipo + "</strong> lideram os erros. Ponto focal: redobrar atenção nesse tipo de lançamento.";
  }
  return nota;
}

function badgeGrav(grav) {
  var g = (grav || "").toUpperCase();
  if (g === "GRAVÍSSIMO") return '<span style="background:#FFEBEE;color:#B71C1C;padding:2px 8px;border-radius:4px;font-size:11px;font-weight:600;">Gravíssimo</span>';
  if (g === "GRAVE")      return '<span style="background:#FFF3E0;color:#E65100;padding:2px 8px;border-radius:4px;font-size:11px;font-weight:600;">Grave</span>';
  if (g === "MÉDIO")      return '<span style="background:#FFF8E1;color:#F57F17;padding:2px 8px;border-radius:4px;font-size:11px;font-weight:600;">Médio</span>';
  if (g === "BAIXO")      return '<span style="background:#E8F5E9;color:#2E7D32;padding:2px 8px;border-radius:4px;font-size:11px;font-weight:600;">Baixo</span>';
  return '<span style="background:#F5F5F5;color:#666;padding:2px 8px;border-radius:4px;font-size:11px;">'+grav+'</span>';
}
function taxaColor(t){ return t>=35?"#B71C1C":t>=20?"#E65100":"#2E7D32"; }
function taxaBg(t)   { return t>=35?"#FFEBEE":t>=20?"#FFF3E0":"#E8F5E9"; }
function barColor(t) { return t>=35?"#E53935":t>=20?"#FB8C00":"#43A047"; }
function contarGravAgente(errosAg, tipoMap, grav) {
  return Object.keys(errosAg).filter(function(t){
    return tipoMap[t] && tipoMap[t].gravidade.toUpperCase()===grav.toUpperCase();
  }).reduce(function(s,t){ return s+errosAg[t]; }, 0);
}

// ============================================================
//  GERADOR DE HTML
// ============================================================
function gerarHTML(hoje, d, periodo, titulo) {
  var maxQtd  = d.tipoArr.length > 0 ? d.tipoArr[0].qtd : 1;
  var maxTaxa = d.agentesArr.length > 0 ? d.agentesArr[0].taxa : 1;

  var rankingRows = d.tipoArr.slice(0,10).map(function(t, i){
    var pct = Math.round((t.qtd/maxQtd)*100);
    return '<tr>' +
      '<td style="padding:8px 12px;color:#888;font-size:12px;">'+(i+1)+'</td>' +
      '<td style="padding:8px 12px;font-size:13px;color:#222;">'+t.tipo+'</td>' +
      '<td style="padding:8px 12px;">'+badgeGrav(t.gravidade)+'</td>' +
      '<td style="padding:8px 12px;font-size:12px;color:#888;">'+t.fase+'</td>' +
      '<td style="padding:8px 12px;min-width:130px;">' +
        '<div style="display:flex;align-items:center;gap:8px;">' +
          '<div style="flex:1;height:5px;background:#EEF0F4;border-radius:3px;">' +
            '<div style="width:'+pct+'%;height:5px;background:#2D6BC4;border-radius:3px;"></div>' +
          '</div>' +
          '<span style="font-weight:600;font-size:13px;color:#1C2B4A;min-width:20px;text-align:right;">'+t.qtd+'</span>' +
        '</div>' +
      '</td>' +
    '</tr>';
  }).join("");

  var agenteBars = d.agentesArr.map(function(ag){
    var pct  = maxTaxa > 0 ? Math.round((ag.taxa/maxTaxa)*100) : 0;
    var nome = ag.nome.replace("BKO - ","");
    return '<div style="margin-bottom:14px;">' +
      '<div style="display:flex;justify-content:space-between;align-items:center;margin-bottom:4px;">' +
        '<span style="font-size:13px;font-weight:500;color:#222;">'+nome+'</span>' +
        '<span style="font-size:12px;color:#888;">'+ag.totalErros+' erros / '+ag.analisados+' analisados &nbsp;' +
          '<strong style="color:'+taxaColor(ag.taxa)+';">'+ag.taxa.toFixed(1)+'%</strong></span>' +
      '</div>' +
      '<div style="height:6px;background:#EEF0F4;border-radius:3px;">' +
        '<div style="width:'+pct+'%;height:6px;background:'+barColor(ag.taxa)+';border-radius:3px;"></div>' +
      '</div>' +
    '</div>';
  }).join("");

  var agentCards = d.agentesArr.map(function(ag){
    var nome       = ag.nome.replace("BKO - ","");
    var taxa       = ag.taxa.toFixed(1);
    var nota       = gerarNota(ag);
    var gravCount  = contarGravAgente(ag.erros, d.tipoMap, "GRAVÍSSIMO");
    var graveCount = contarGravAgente(ag.erros, d.tipoMap, "GRAVE");
    var baixoCount = contarGravAgente(ag.erros, d.tipoMap, "BAIXO") + contarGravAgente(ag.erros, d.tipoMap, "MÉDIO");
    var alertLabel = ag.taxa >= 35 ? "Atenção imediata" : ag.taxa >= 20 ? "Atenção" : "Monitorar";

    var topErros = Object.keys(ag.erros).map(function(t){
      return { tipo: t, qtd: ag.erros[t] };
    }).sort(function(a,b){ return b.qtd-a.qtd; }).slice(0,5);

    var erroRows = topErros.map(function(e){
      return '<div style="display:flex;justify-content:space-between;align-items:center;padding:7px 0;border-bottom:1px solid #F5F5F5;">' +
        '<span style="font-size:13px;color:#333;">'+e.tipo+'</span>' +
        '<span style="font-size:12px;color:#888;white-space:nowrap;margin-left:12px;">'+e.qtd+' caso'+(e.qtd>1?'s':'')+'</span>' +
      '</div>';
    }).join("");

    var naExp = Object.keys(ag.erros).filter(function(t){ return d.tipoMap[t] && d.tipoMap[t].fase==="EXPORTAÇÃO"; })
                  .reduce(function(s,t){ return s+ag.erros[t]; }, 0);
    var naAna = Object.keys(ag.erros).filter(function(t){ return d.tipoMap[t] && d.tipoMap[t].fase==="ANÁLISE"; })
                  .reduce(function(s,t){ return s+ag.erros[t]; }, 0);
    var faseText = [];
    if (naExp > 0) faseText.push(naExp+" na exportação");
    if (naAna > 0) faseText.push(naAna+" na análise");

    var badgesGrav = "";
    if (gravCount  > 0) badgesGrav += '<span style="background:#FFEBEE;color:#B71C1C;padding:3px 10px;border-radius:4px;font-size:12px;font-weight:600;">'+gravCount+' gravíssimo'+(gravCount>1?'s':'')+'</span> ';
    if (graveCount > 0) badgesGrav += '<span style="background:#FFF3E0;color:#E65100;padding:3px 10px;border-radius:4px;font-size:12px;font-weight:600;">'+graveCount+' grave'+(graveCount>1?'s':'')+'</span> ';
    if (baixoCount > 0) badgesGrav += '<span style="background:#E8F5E9;color:#2E7D32;padding:3px 10px;border-radius:4px;font-size:12px;font-weight:600;">'+baixoCount+' baixo'+(baixoCount>1?'s':'')+'</span>';

    return '<div style="background:#fff;border:1px solid #E8E8E8;border-radius:12px;padding:20px 22px;margin-bottom:16px;">' +
      '<div style="display:flex;justify-content:space-between;align-items:flex-start;margin-bottom:14px;">' +
        '<div>' +
          '<div style="font-size:17px;font-weight:600;color:#1a1a1a;">'+nome+'</div>' +
          '<div style="font-size:12px;color:#999;margin-top:3px;">'+alertLabel+' · '+ag.analisados+' registros analisados</div>' +
        '</div>' +
        '<div style="background:'+taxaBg(ag.taxa)+';color:'+taxaColor(ag.taxa)+';padding:5px 13px;border-radius:20px;font-size:13px;font-weight:600;">'+taxa+'% de erro</div>' +
      '</div>' +
      (badgesGrav ? '<div style="display:flex;gap:8px;flex-wrap:wrap;margin-bottom:14px;">'+badgesGrav+'</div>' : '') +
      '<div style="font-size:10px;font-weight:600;letter-spacing:0.8px;color:#999;text-transform:uppercase;margin-bottom:8px;">Principais ocorrências</div>' +
      erroRows +
      (faseText.length > 0 ? '<div style="font-size:12px;color:#999;margin-top:10px;">Fase: '+faseText.join(' · ')+'</div>' : '') +
      '<div style="background:#FFF8F8;border-left:3px solid #E53935;border-radius:0 6px 6px 0;padding:12px 14px;margin-top:14px;">' +
        '<div style="font-size:12px;font-weight:600;color:#B71C1C;margin-bottom:5px;">Nota para sinalização</div>' +
        '<div style="font-size:12px;color:#555;line-height:1.6;">'+nota+'</div>' +
      '</div>' +
    '</div>';
  }).join("");

  return '<!DOCTYPE html><html lang="pt-BR"><head><meta charset="UTF-8">' +
  '<meta name="viewport" content="width=device-width,initial-scale=1">' +
  '<title>'+titulo+' — Qualidade Backoffice</title>' +
  '<style>body{font-family:Arial,sans-serif;background:#F0F2F5;color:#1a1a1a;margin:0;padding:32px 20px;}' +
  '.container{max-width:820px;margin:0 auto;}table{width:100%;border-collapse:collapse;}' +
  'th{padding:9px 12px;text-align:left;font-size:11px;font-weight:600;color:#888;border-bottom:1px solid #EBEBEB;' +
  'background:#FAFAFA;text-transform:uppercase;letter-spacing:.4px;}' +
  'tr:last-child td{border-bottom:none;}</style></head><body>' +
  '<div class="container">' +

  '<div style="background:#1C2B4A;border-radius:14px;padding:26px 30px;margin-bottom:18px;display:flex;justify-content:space-between;align-items:center;">' +
    '<div>' +
      '<div style="font-size:11px;font-weight:600;letter-spacing:1.2px;text-transform:uppercase;color:#7BA7D8;margin-bottom:6px;">Backoffice · '+titulo+'</div>' +
      '<div style="font-size:22px;font-weight:700;color:#fff;">Análise de Erros — '+mesAnoAtual()+'</div>' +
      '<div style="font-size:12px;color:#7BA7D8;margin-top:5px;">'+d.totalAnalisado+' registros analisados · '+d.nAgentes+' agentes · Período: '+periodo+'</div>' +
    '</div>' +
    '<div style="text-align:right;">' +
      '<div style="font-size:40px;font-weight:700;color:#fff;line-height:1;">'+d.totalErros+'</div>' +
      '<div style="font-size:12px;color:#7BA7D8;margin-top:2px;">erros identificados</div>' +
    '</div>' +
  '</div>' +

  '<div style="display:grid;grid-template-columns:repeat(4,1fr);gap:12px;margin-bottom:18px;">' +
    '<div style="background:#fff;border:1px solid #E4E8EE;border-radius:10px;padding:16px 18px;"><div style="font-size:11px;color:#8A93A2;font-weight:600;margin-bottom:6px;text-transform:uppercase;letter-spacing:.5px;">Total analisado</div><div style="font-size:26px;font-weight:700;color:#1C2B4A;">'+d.totalAnalisado+'</div></div>' +
    '<div style="background:#fff;border:1px solid #E4E8EE;border-radius:10px;padding:16px 18px;"><div style="font-size:11px;color:#8A93A2;font-weight:600;margin-bottom:6px;text-transform:uppercase;letter-spacing:.5px;">Com erro</div><div style="font-size:26px;font-weight:700;color:#1C2B4A;">'+d.totalErros+'</div></div>' +
    '<div style="background:#fff;border:1px solid #E4E8EE;border-radius:10px;padding:16px 18px;"><div style="font-size:11px;color:#8A93A2;font-weight:600;margin-bottom:6px;text-transform:uppercase;letter-spacing:.5px;">Taxa geral</div><div style="font-size:26px;font-weight:700;color:#C0392B;">'+d.taxaGeral+'%</div></div>' +
    '<div style="background:#fff;border:1px solid #E4E8EE;border-radius:10px;padding:16px 18px;"><div style="font-size:11px;color:#8A93A2;font-weight:600;margin-bottom:6px;text-transform:uppercase;letter-spacing:.5px;">Gravíssimos</div><div style="font-size:26px;font-weight:700;color:#C0392B;">'+d.gravissimos+'</div></div>' +
  '</div>' +

  '<div style="background:#fff;border:1px solid #E4E8EE;border-radius:10px;padding:20px 22px;margin-bottom:18px;">' +
    '<div style="font-size:11px;font-weight:600;letter-spacing:.8px;text-transform:uppercase;color:#999;margin-bottom:16px;">Taxa de erro por agente</div>' +
    agenteBars +
  '</div>' +

  '<div style="background:#fff;border:1px solid #E4E8EE;border-radius:10px;overflow:hidden;margin-bottom:24px;">' +
    '<div style="background:#1C2B4A;padding:12px 18px;font-size:12px;font-weight:600;color:#fff;letter-spacing:.5px;text-transform:uppercase;">Ranking de erros — Top 10</div>' +
    '<table><thead><tr>' +
      '<th style="width:4%">#</th><th style="width:48%">Tipo de erro</th><th>Gravidade</th><th>Fase</th><th style="width:22%">Qtd</th>' +
    '</tr></thead><tbody>'+rankingRows+'</tbody></table>' +
  '</div>' +

  '<div style="font-size:11px;font-weight:600;letter-spacing:.8px;text-transform:uppercase;color:#999;margin-bottom:14px;">Detalhamento por agente</div>' +
  agentCards +

  '<div style="text-align:center;font-size:11px;color:#AAB0BB;margin-top:24px;padding-bottom:8px;">' +
    'Sistema de Qualidade · Backoffice · Gerado em '+hoje +
  '</div>' +
  '</div></body></html>';
}

// ── MENU ─────────────────────────────────────────────────────
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu("Relatório de Qualidade")
    .addItem("Gerar e enviar relatório", "gerarRelatorio")
    .addToUi();
}
