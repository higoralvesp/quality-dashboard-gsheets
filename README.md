# quality-dashboard-gsheets
Dashboard interativo de controle de qualidade para equipes de backoffice, com relatório automático por e-mail via Google Apps Script
# 📊 Dashboard de Controle de Qualidade — Backoffice

Sistema completo de acompanhamento de qualidade operacional para equipes de análise documental, desenvolvido em Google Sheets com automação via Google Apps Script.

## 🎯 Problema resolvido

Equipes de backoffice que revisam grandes volumes de registros diariamente precisam de visibilidade sobre:
- Quantos registros cada analista processou
- Quais erros estão ocorrendo com mais frequência
- Como a taxa de erro evolui semana a semana
- Quem precisa de atenção imediata

Este sistema centraliza tudo isso em um dashboard interativo e envia relatórios automáticos para a supervisão.

## ✅ Funcionalidades

- **Dashboard interativo** com dropdown por analista
- **Visão geral do mês** com taxa de erro por agente
- **Produção diária** com destaque automático para finais de semana e feriados
- **Evolução semanal** da taxa de erro
- **Ranking de erros** por tipo e gravidade
- **Relatório HTML automático** enviado por e-mail para supervisão (consolidado mensal + diagnóstico do período)
- **Texto formatado** gerado automaticamente para comunicação via grupo

## 🛠️ Tecnologias

- Google Sheets (fórmulas avançadas: COUNTIFS, QUERY, PROCV, formatação condicional)
- Google Apps Script (JavaScript)
- Gmail API via Apps Script
- HTML/CSS inline para relatórios por e-mail

## 📁 Estrutura
├── relatorio_bko.gs       # Script principal de geração e envio de relatórios

└── README.md

## 📸 Preview

*(em breve — print do dashboard e do relatório gerado)*

## 💡 Como funciona o relatório automático

O script detecta automaticamente o dia da semana e define o período de análise:
- **Segunda-feira:** puxa dados da semana anterior
- **Terça a sexta:** puxa da semana atual até o dia anterior  
- **Sábado:** puxa a semana completa

Gera dois relatórios HTML e envia por e-mail, além de um texto formatado para o grupo da equipe.

## 📌 Contexto

Projeto desenvolvido a partir de necessidade real em ambiente corporativo de backoffice. Dados utilizados neste repositório são fictícios.
