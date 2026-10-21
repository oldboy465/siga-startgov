# 📊Painel de Gestão de Processos SIGA / StartGov | UEMA

<div align="center">

![UEMA](https://img.shields.io/badge/UEMA-Universidade%20Estadual%20do%20Maranh%C3%A3o-003366?style=for-the-badge&logo=academia&logoColor=white)
![SIGA](https://img.shields.io/badge/SIGA-Sistema%20Integrado-c59b27?style=for-the-badge)
![StartGov](https://img.shields.io/badge/StartGov-Plataforma-1e5a96?style=for-the-badge)
![Versão](https://img.shields.io/badge/Vers%C3%A3o-1.0.0-success?style=for-the-badge)
![Status](https://img.shields.io/badge/Status-Produ%C3%A7%C3%A3o-brightgreen?style=for-the-badge)
![Licença](https://img.shields.io/badge/Licen%C3%A7a-Institucional-blue?style=for-the-badge)

### 🏛️ Painel de Acompanhamento de Processos Licitatórios e Orçamentários

**Uma aplicação web SPA (Single Page Application) 100% client-side, sem backend, sem banco de dados, sem dependências de servidor — roda inteiramente no navegador.**

</div>

---

## 📑 Índice

- [✨ Visão Geral](#-visão-geral)
- [🎯 Proposta e Objetivo](#-proposta-e-objetivo)
- [🚀 Demonstração Visual](#-demonstração-visual)
- [🧱 Stack Tecnológico](#-stack-tecnológico)
- [📦 Instalação e Uso](#-instalação-e-uso)
- [🗂️ Estrutura do Projeto](#️-estrutura-do-projeto)
- [🎨 Design System e Paleta de Cores](#-design-system-e-paleta-de-cores)
- [🏗️ Arquitetura da Aplicação](#️-arquitetura-da-aplicação)
- [🧠 Estado Global (`state`)](#-estado-global-state)
- [🔧 Módulo: Importação de Planilhas](#-módulo-importação-de-planilhas)
- [📊 Módulo: Dashboard (KPIs + Gráficos)](#-módulo-dashboard-kpis--gráficos)
- [📋 Módulo: Tabela Analítica](#-módulo-tabela-analítica)
- [🎛️ Módulo: Gerenciador de Colunas](#️-módulo-gerenciador-de-colunas)
- [🔍 Módulo: Filtros Dinâmicos Multi-Seleção](#-módulo-filtros-dinâmicos-multi-seleção)
- [✏️ Módulo: Edição Inline de Células](#️-módulo-edição-inline-de-células)
- [🖱️ Módulo: Drag & Drop de Colunas](#️-módulo-drag--drop-de-colunas)
- [📤 Módulo: Exportação (XLSX / PDF / WhatsApp)](#-módulo-exportação-xlsx--pdf--whatsapp)
- [🛡️ Módulo: Auditoria de Erros](#️-módulo-auditoria-de-erros)
- [📅 Sistema de Datas](#-sistema-de-datas)
- [💰 Sistema Monetário](#-sistema-monetário)
- [🎬 Animações e Efeitos Visuais](#-animações-e-efeitos-visuais)
- [🧩 API Pública (`window.app`)](#-api-pública-windowapp)
- [📐 Regras de Negócio e Convenções](#-regras-de-negócio-e-convenções)
- [🧪 Fluxo Completo do Usuário](#-fluxo-completo-do-usuário)
- [🎯 Casos de Uso Reais](#-casos-de-uso-reais)
- [🔒 Segurança e Privacidade](#-segurança-e-privacidade)
- [🌐 Compatibilidade](#-compatibilidade)
- [🐛 Limitações Conhecidas](#-limitações-conhecidas)
- [🗺️ Roadmap Futuro](#️-roadmap-futuro)
- [🤝 Contribuição](#-contribuição)
- [📜 Créditos e Licença](#-créditos-e-licença)

---

## ✨ Visão Geral

O **Painel de Gestão de Processos SIGA / StartGov** é uma aplicação web institucional desenvolvida para a **Universidade Estadual do Maranhão (UEMA)**, projetada para transformar planilhas Excel brutas de processos licitatórios e orçamentários em **dashboards executivos interativos**, **tabelas analíticas editáveis** e **relatórios PDF de alto padrão**.

> 💡 **Filosofia do projeto:** "Do Excel bruto ao relatório executivo em menos de 30 segundos — sem servidor, sem banco, sem complicação."

### 🌟 Destaques Principais

| # | Recurso | Descrição |
|---|---------|-----------|
| 🎯 | **100% Client-Side** | Nenhum dado sai do navegador. Nenhum backend. Nenhum banco. |
| 📥 | **Importação Inteligente** | Detecção automática da linha de cabeçalho, ignorando linhas prévias |
| 📊 | **6 KPIs Executivos** | Processos, montante, autorizados, pendentes, SIGA, contratados |
| 📈 | **4 Gráficos Interativos** | Modalidade, Fonte, Localização, Status (Chart.js 4.4.1) |
| 🔍 | **Filtros Multi-Seleção** | Dropdown com busca textual + contadores por opção |
| ✏️ | **Edição Inline** | Duplo clique em qualquer célula para editar |
| 🖱️ | **Drag & Drop** | Reordenação de colunas por arrastar |
| 📤 | **4 Modos de Exportação** | XLSX, PDF (A4 Paisagem), WhatsApp, Auditoria |
| 🎨 | **Design Institucional** | Paleta oficial UEMA com gradientes e sombras |
| ♿ | **Acessibilidade** | Contraste WCAG AA em cabeçalhos e textos |

---

## 🎯 Proposta e Objetivo

### O Problema

Setores de licitação e orçamento da UEMA recebem, diariamente, **planilhas Excel** contendo dezenas a centenas de processos com estrutura irregular:

- Cabeçalhos começando na linha 3, 4 ou 5 (linhas de título antes)
- Valores monetários em formatos mistos (`R$ 1.234,56`, `1234.56`, `-`)
- Datas como texto (`10/03/2026`), número Excel serial ou ISO
- Erros de fórmula (`#VALUE!`, `#REF!`, `#NAME?`) silenciosos
- Colunas vazias, renomeadas ou com espaços extras
- Necessidade de relatórios de status para a reitoria

### A Solução

Este painel entrega:

1. **Parsing resiliente** da planilha com preview visual e escolha da linha de cabeçalho
2. **KPIs e gráficos** gerados dinamicamente a partir das colunas detectadas
3. **Tabela interativa** com ordenação, paginação, edição e filtros múltiplos
4. **Exportação** para PDF executivo com seleção de colunas (proteção anti-quebra)
5. **Mensagem WhatsApp** pré-formatada com resumo executivo e detalhamento
6. **Auditoria de erros** de fórmula da planilha original

---

## 🚀 Demonstração Visual

```
┌──────────────────────────────────────────────────────────────────────────┐
│  🏛️ PAINEL DE ACOMPANHAMENTO DE PROCESSOS                                │
│  [SIGA & STARTGOV] Universidade Estadual do Maranhão | Gestão de Licitações│
│  ─────────────────────────────────────────────────────────────────────── │
│  [📥 Importar Excel] [💬 Copiar WhatsApp] [📗 Exportar XLSX] [📕 PDF] [⚠️]│
├──────────────────────────────────────────────────────────────────────────┤
│  [📊 Indicadores & Gráficos] [📋 Tabela Analítica] [🎛️ Gerenciar Colunas] │
├──────────────────────────────────────────────────────────────────────────┤
│  Arquivo: processos_2026.xlsx | Planilha: SIGA 2026 | Linhas ignoradas: 3 │
│  Registros: 87 | Colunas Ativas: 14                                       │
├──────────────────────────────────────────────────────────────────────────┤
│  📁 Processos: 87   💰 Montante: R$ 12.4M   ✅ Autorizados: 42 (48.3%)   │
│  ⏳ Pendentes: 31   🔗 SIGA: 79   📄 Contratados: 22 (25.3%)             │
├──────────────────────────────────────────────────────────────────────────┤
│  📊 Gráfico Modalidade    │  🥯 Gráfico Fonte de Recursos               │
│  📍 Top 8 Localizações    │  🥧 Gráfico Status                          │
└──────────────────────────────────────────────────────────────────────────┘
```

---

## 🧱 Stack Tecnológico

| Camada | Tecnologia | Versão | CDN |
|--------|-----------|--------|-----|
| **Markup** | HTML5 Semântico | — | Nativo |
| **Estilo** | CSS3 (Grid, Flexbox, Variáveis CSS, Animações) | — | Inline `<style>` |
| **Lógica** | JavaScript ES6+ (IIFE, Closures, Set/Map) | — | Inline `<script>` |
| **Planilhas** | SheetJS (xlsx) | `0.18.5` | `cdn.jsdelivr.net` |
| **Gráficos** | Chart.js | `4.4.1` | `cdn.jsdelivr.net` |
| **PDF Core** | jsPDF | `2.5.1` | `cdnjs.cloudflare.com` |
| **PDF Tabelas** | jspdf-autotable | `3.8.2` | `cdnjs.cloudflare.com` |
| **Ícones** | Font Awesome | `6.4.2` | `cdnjs.cloudflare.com` |

### 🔗 URLs das Dependências (Offline-Ready)

```html
<script src="https://cdn.jsdelivr.net/npm/xlsx@0.18.5/dist/xlsx.full.min.js"></script>
<script src="https://cdn.jsdelivr.net/npm/chart.js@4.4.1/dist/chart.umd.min.js"></script>
<script src="https://cdnjs.cloudflare.com/ajax/libs/jspdf/2.5.1/jspdf.umd.min.js"></script>
<script src="https://cdnjs.cloudflare.com/ajax/libs/jspdf-autotable/3.8.2/jspdf.plugin.autotable.min.js"></script>
<link rel="stylesheet" href="https://cdnjs.cloudflare.com/ajax/libs/font-awesome/6.4.2/css/all.min.css">
```

---

## 📦 Instalação e Uso

### Instalação Mínima (Zero Config)

```bash
# 1. Clone ou baixe o arquivo index.html
# 2. Abra diretamente no navegador (duplo clique)

# Ou, se preferir servir localmente:
python -m http.server 8000
# Acesse http://localhost:8000
```

### Requisitos

- ✅ Navegador moderno (Chrome 90+, Firefox 88+, Edge 90+, Safari 14+)
- ✅ Conexão com a internet (apenas para os CDNs na primeira carga)
- ✅ Planilha `.xlsx` ou `.xls`
- ❌ **Não requer:** Node.js, NPM, backend, banco de dados, autenticação

### Primeiros Passos

1. **Importar planilha** → Clique em `📥 IMPORTAR EXCEL` no cabeçalho
2. **Configurar** → Selecione a aba, ajuste quantas linhas ignorar, visualize o preview
3. **Processar** → Clique em `✅ Processar e Importar Planilha`
4. **Explorar** → Navegue entre as abas, aplique filtros, edite, exporte

---

## 🗂️ Estrutura do Projeto

```
📦 painel-siga-uema/
├── 📄 index.html          ← Arquivo único (single file app)
│   ├── <head>
│   │   ├── Metadados & viewport
│   │   ├── 5 CDNs externos
│   │   └── <style> (~900 linhas de CSS)
│   ├── <body>
│   │   ├── <header> — Cabeçalho institucional + botões
│   │   ├── <nav> — 3 abas SPA
│   │   ├── <main> — Container das 3 abas
│   │   │   ├── #tab-dashboard
│   │   │   ├── #tab-table
│   │   │   └── #tab-columns
│   │   ├── 3 Modais (importação, PDF, erros)
│   │   ├── Toast WhatsApp
│   │   └── <footer> institucional
│   └── <script> (~1400 linhas JS)
└── 📘 README.md           ← Este arquivo
```

### 🏛️ Anatomia de um Único Arquivo

| Seção | Localização | Linhas Aprox. |
|-------|-------------|---------------|
| Meta & CDN | `<head>` | 1–20 |
| Design System (CSS) | `<style>` | 21–900 |
| Cabeçalho | `<header.institutional-header>` | 901–940 |
| Navegação | `<nav.app-nav>` | 941–950 |
| Metadados Banner | `.meta-banner` | 960–980 |
| Dashboard | `#tab-dashboard` | 981–1100 |
| Tabela | `#tab-table` | 1101–1250 |
| Gerenciador Colunas | `#tab-columns` | 1251–1300 |
| Modais | 3 blocos `.modal-overlay` | 1301–1450 |
| Toast | `#whatsappToast` | 1451–1460 |
| Motor JS | `<script>` | 1461–2900 |

---

## 🎨 Design System e Paleta de Cores

### Variáveis CSS (`:root`)

| Variável | Valor Hex | Uso |
|----------|-----------|-----|
| `--uema-blue` | `#003366` | Cor primária institucional |
| `--uema-dark-blue` | `#002244` | Cabeçalhos, títulos |
| `--uema-light-blue` | `#1e5a96` | Hover, links |
| `--uema-accent` | `#d9534f` | Destaques |
| `--uema-gold` | `#c59b27` | Faixa dourada do header |
| `--whatsapp-green` | `#25d366` | Botão WhatsApp |
| `--whatsapp-dark` | `#128c7e` | Hover WhatsApp |
| `--bg-body` | `#f4f6f9` | Fundo geral |
| `--bg-card` | `#ffffff` | Cartões |
| `--text-main` | `#2b303a` | Texto principal |
| `--text-muted` | `#64748b` | Texto secundário |
| `--border-color` | `#dce2e6` | Bordas |
| `--success-color` | `#28a745` | Sucesso |
| `--warning-color` | `#ffc107` | Alerta |
| `--danger-color` | `#dc3545` | Perigo |
| `--info-color` | `#17a2b8` | Informação |

### Sombras (Elevation System)

| Nome | Valor CSS | Nível |
|------|-----------|-------|
| `--shadow-sm` | `0 2px 5px rgba(0,0,0,0.05)` | Nível 1 |
| `--shadow-md` | `0 4px 12px rgba(0,0,0,0.08)` | Nível 2 |
| `--shadow-lg` | `0 8px 24px rgba(0,0,0,0.12)` | Nível 3 |

### 🎬 Sistema de Animações Completo

O projeto entrega uma experiência **rica em microinterações e animações CSS**, sempre com foco em feedback visual e percepção de fluidez:

#### 1️⃣ Animações de Entrada

- **`toastSlideUp`** — Toast WhatsApp desliza de baixo para cima com fade-in
  ```css
  @keyframes toastSlideUp {
      from { transform: translateY(20px); opacity: 0; }
      to   { transform: translateY(0);    opacity: 1; }
  }
  animation: toastSlideUp 0.3s ease-out;
  ```

#### 2️⃣ Animações de Hover (Botões e Cards)

- `btn`, `btn-primary`, `btn-success`, `btn-whatsapp`, `btn-warning`, `btn-danger` → **transição de cor em 0.2s**
- `multi-filter-btn:hover` → **borda azul iluminada**
- `multi-filter-item:hover` → **fundo cinza-claro (`#f1f5f9`)**
- `siga-table tbody tr:hover` → **fundo azul-tint (`#f1f6fc`)**
- `column-dropdown-item:hover` → **fundo azul-tint (`#f0f6ff`)**
- `editable-cell:hover` → **fundo azul + sombra inset (`box-shadow: inset 0 0 0 1px #90caf9`)**
- `nav-tab:hover` → **cor azul com border-bottom**

#### 3️⃣ Animações de Estado (Ativo / Filtrado)

- `.nav-tab.active` → **border-bottom azul de 3px + fundo `#f8fafc`**
- `.multi-filter-btn.is-filtered` → **fundo `#eff6ff` + negrito**
- `.active-filter-pill` → **pill azul (`#e0f2fe`)**
- `.active-filter-pill.has-filters` → **pill amarela (`#fef3c7`)** com borda
- `.column-dropdown-menu.open` → **display block**
- `.multi-filter-dropdown-menu.open` → **display block**

#### 4️⃣ Animações de Drag & Drop

- `th.drag-over` → **background gold (`--uema-gold`) + texto preto** (feedback instantâneo de drop válido)

#### 5️⃣ Transições Globais

- `button`, `a`, `tr`, `td`, `th` → `transition: all 0.2s ease`
- `table-filter-section` → `transition: all 0.3s ease`
- `multi-filter-btn` → `transition: all 0.2s ease`

#### 6️⃣ Efeitos Visuais Adicionais

- **Gradiente no header:** `linear-gradient(135deg, #002244, #003366)`
- **Faixa dourada:** `border-bottom: 3px solid var(--uema-gold)`
- **Bordas coloridas nos KPIs:** `border-left: 4px solid {cor específica}`
- **Ícones coloridos por KPI:**
  - 📁 Azul UEMA
  - 💰 Verde (`#28a745`)
  - ✅ Ciano (`#17a2b8`)
  - ⏳ Amarelo (`#d39e00`)
  - 🔗 Roxo (`#6f42c1`)
  - 📄 Verde-água (`#20c997`)
- **Tabela preview importação:**
  - Linha ignorada → `#fee2e2` (vermelho claro)
  - Linha cabeçalho → `#dbeafe` (azul claro)
  - Linha dados → `#f0fdf4` (verde claro)

#### 7️⃣ Animações Recomendadas (Extensões Futuras Sugeridas)

> 💡 **Sugestões de upgrade estético** (ainda não implementadas):
> - `@keyframes pulse` para KPIs com valores críticos
> - `@keyframes shimmer` para skeleton loading durante importação
> - `@keyframes float` para ícones do header
> - `@keyframes gradientShift` para background animado do header
> - `@keyframes numberRoll` para contagem incremental dos KPIs
> - `@keyframes wiggle` para feedback de erro em validação
> - Transições `view-transition-name` para navegação entre abas (Chrome 111+)
> - `prefers-reduced-motion` para acessibilidade

---

## 🏗️ Arquitetura da Aplicação

### Padrão: Module Revealing (IIFE)

Toda a aplicação está encapsulada em uma **IIFE (Immediately Invoked Function Expression)** que expõe apenas os métodos necessários via `window.app`:

```javascript
window.app = (function() {
    'use strict';
    
    let state = { /* ... */ };  // Estado privado
    
    // ... ~40 funções privadas ...
    
    return {
        init, switchTab, reopenSheetModal,
        // ... API pública exposta ...
    };
})();

window.addEventListener('DOMContentLoaded', () => window.app.init());
```

### 🎯 Vantagens do Padrão

| Aspecto | Benefício |
|---------|-----------|
| **Encapsulamento** | `state` não pode ser modificado externamente |
| **Sem conflito global** | Apenas `window.app` é exposto |
| **Testável** | Métodos expostos podem ser testados |
| **Sem bundler** | Roda puro no navegador |

### 🔄 Fluxo de Dados (Unidirecional)

```
┌─────────────┐
│   Arquivo   │
│    XLSX     │
└──────┬──────┘
       │ FileReader → XLSX.read
       ▼
┌─────────────────┐
│  rawMatrix[][]  │  (matriz bruta)
└──────┬──────────┘
       │ confirmSheetImport()
       ▼
┌─────────────────┐
│  state.rows[]   │  (registros parseados)
└──────┬──────────┘
       │ applyFilters()
       ▼
┌────────────────────┐
│ state.filteredRows │
└──────┬─────────────┘
       │ renderKPIs() / renderCharts() / renderTable()
       ▼
┌──────────────┐
│  DOM Visual  │
└──────────────┘
```

---

## 🧠 Estado Global (`state`)

O objeto `state` é a **única fonte de verdade** da aplicação:

```javascript
let state = {
    // ─── IMPORTAÇÃO ───
    workbook: null,                    // Objeto SheetJS completo
    fileName: "",                      // Nome do arquivo importado
    currentSheet: "",                  // Nome da aba ativa
    ignoredRowsCount: 3,               // Linhas ignoradas antes do header
    rawMatrix: [],                     // Matriz bruta da planilha
    
    // ─── COLUNAS ───
    headers: [],                       // Colunas ativas (ordem de exibição)
    allAvailableHeaders: [],           // Todas as colunas detectadas
    visibleHeaders: [],                // Colunas marcadas para exibição
    
    // ─── DADOS ───
    rows: [],                          // Base completa de registros
    filteredRows: [],                  // Registros após filtros
    errors: [],                        // Log de auditoria
    
    // ─── FILTROS ───
    filters: {
        search: "",                    // Busca global
        multiSelect: {}                // { colName: Set() }
    },
    
    // ─── UI ───
    charts: {},                        // Instâncias Chart.js por canvas
    pagination: {
        currentPage: 1,
        recordsPerPage: 30,
        sortCol: null,
        sortDir: 'asc'
    },
    dragSrcColIdx: null,               // Índice da coluna arrastada
    pdfSelectedCols: new Set()         // Colunas selecionadas para PDF
};
```

### 🏷️ Metadados por Registro

Cada registro em `state.rows` é um objeto com:

| Propriedade | Tipo | Descrição |
|-------------|------|-----------|
| `_id` | `number` | ID único = número da linha original no Excel |
| `[headerName]` | `any` | Valor bruto da célula |
| `_num_[headerName]` | `number` | Valor numérico pré-calculado (para colunas com "VALOR") |

---

## 🔧 Módulo: Importação de Planilhas

### Funções Envolvidas

| Função | Descrição |
|--------|-----------|
| `handleFileInput(e)` | Captura o arquivo via `<input type="file">` |
| `showConfigModal(workbook)` | Abre modal com seletores de aba e linha |
| `handleModalSheetChange(sheetName)` | Detecta automaticamente a linha de cabeçalho |
| `updateSheetPreview()` | Renderiza preview colorido das 10 primeiras linhas |
| `setIgnoreRowsCount(idx)` | Permite clicar em uma linha para torná-la cabeçalho |
| `confirmSheetImport()` | Executa o parsing final e popula o `state` |
| `reopenSheetModal()` | Reabre o modal para reconfiguração |

### 🔍 Detecção Automática de Cabeçalho

O algoritmo percorre as **10 primeiras linhas** e identifica a primeira que contenha:

- **≥ 3 células preenchidas** ✅
- **E algum dos termos-chave:** `PROCESSO`, `MODALIDADE`, `OBJETO` ✅

```javascript
const filledCount = row.filter(c => c !== null && String(c).trim() !== "").length;
if (filledCount >= 3 && strRow.some(c => 
    c.includes("PROCESSO") || c.includes("MODALIDADE") || c.includes("OBJETO")
)) {
    autoIgnore = r;
    break;
}
```

### 🎨 Preview Visual Colorido

O modal exibe uma tabela com **classes semânticas**:

| Classe CSS | Cor | Significado |
|------------|-----|-------------|
| `.row-ignored` | Vermelho `#fee2e2` | Será ignorada |
| `.row-header` | Azul `#dbeafe` | Cabeçalho selecionado |
| `.row-data` | Verde `#f0fdf4` | Dados válidos |

### 🛡️ Validação de Fórmulas

Durante o parsing, células com `#VALUE!`, `#NAME?` ou `#REF!` são registradas em `state.errors`:

```javascript
if (String(val).includes("#VALUE!") || 
    String(val).includes("#NAME?") || 
    String(val).includes("#REF!")) {
    errors.push({ 
        line: r + 1, 
        col: h, 
        val: val, 
        error: "Erro de fórmula na planilha de origem" 
    });
}
```

### 💰 Pré-cálculo de Valores Monetários

Para toda coluna cujo nome contenha "VALOR", o sistema armazena paralelamente o valor numérico:

```javascript
if (h.toUpperCase().includes("VALOR")) {
    let num = 0;
    if (typeof val === 'number') num = val;
    else if (val) {
        const clean = String(val).replace(/[R$\s]/g, '').replace(',', '.');
        const parsed = parseFloat(clean);
        num = !isNaN(parsed) ? parsed : 0;
    }
    rowObj[`_num_${h}`] = num;
}
```

---

## 📊 Módulo: Dashboard (KPIs + Gráficos)

### 📁 6 KPIs Executivos

| # | KPI | Ícone | Cor Borda | Fonte do Cálculo |
|---|-----|-------|-----------|------------------|
| 1 | **Processos Ativos** | `fa-folder-tree` | Azul UEMA | `state.filteredRows.length` |
| 2 | **Montante Total** | `fa-money-bill-wave` | Verde `#28a745` | Soma de `_num_*` |
| 3 | **Autorizados** | `fa-circle-check` | Ciano `#17a2b8` | Texto contém "AUTORIZADO" |
| 4 | **Aguard. Autorização** | `fa-clock-rotate-left` | Amarelo `#ffc107` | Texto contém "AGUARD" |
| 5 | **Cadastrados no SIGA** | `fa-network-wired` | Roxo `#6f42c1` | Coluna SIGA preenchida |
| 6 | **Contratados / Assinados** | `fa-file-contract` | Verde-água `#20c997` | Contém "CONTRATADO/FINALIZADO/ASSINADO" |

### 🧮 Função `renderKPIs()`

Executada após **toda alteração de filtros**, garantindo que os KPIs reflitam o subconjunto filtrado:

```javascript
function renderKPIs() {
    const total = state.filteredRows.length;
    let sum = 0, authorized = 0, pending = 0, siga = 0, contracted = 0;
    
    state.filteredRows.forEach(r => {
        // Soma valores
        for (let k in r) {
            if (k.startsWith('_num_')) sum += (r[k] || 0);
        }
        
        // Análise textual
        const textRow = Object.values(r).join(' ').toUpperCase();
        if (textRow.includes("AUTORIZADO")) authorized++;
        if (textRow.includes("AGUARD")) pending++;
        if (textRow.includes("CONTRATADO") || 
            textRow.includes("FINALIZADO") || 
            textRow.includes("ASSINADO")) contracted++;
        
        // SIGA preenchido
        const sigaField = Object.entries(r).find(([k]) => k.toUpperCase().includes("SIGA"));
        if (sigaField && String(sigaField[1]).trim() !== "" && String(sigaField[1]).trim() !== "_") {
            siga++;
        }
    });
    
    const avg = total > 0 ? sum / total : 0;
    
    // Atualiza DOM
    document.getElementById('kpiCount').innerText = total;
    document.getElementById('kpiSum').innerText = formatCurrency(sum);
    document.getElementById('kpiAvg').innerText = `Média: ${formatCurrency(avg)}`;
    // ... etc
}
```

### 📈 4 Gráficos Chart.js

| # | Canvas ID | Tipo | Eixo | Coluna Detectada |
|---|-----------|------|------|------------------|
| 1 | `chartModalidade` | Bar | Processos | Contém "MODALIDADE" |
| 2 | `chartFonte` | Doughnut | R$ | Contém "FONTE" |
| 3 | `chartLocalizacao` | Bar (top 8) | Processos | Contém "LOCALIZA" |
| 4 | `chartStatus` | Pie | Processos | Contém "STATUS" |

#### 🎨 Paletas de Cores

**Doughnut (Fonte):**
```javascript
backgroundColor: ['#003366', '#28a745', '#ffc107', '#17a2b8', '#6f42c1', '#fd7e14']
```

**Pie (Status):**
```javascript
backgroundColor: ['#28a745', '#ffc107', '#dc3545', '#17a2b8', '#6c757d']
```

#### ♻️ Ciclo de Vida dos Gráficos

```javascript
if (state.charts[canvasId]) state.charts[canvasId].destroy();  // Limpa instância anterior
// ... cria novo Chart ...
```

Isso evita o clássico erro de **"Canvas is already in use"** do Chart.js.

#### 💡 Tooltips Formatados

```javascript
tooltip: { 
    callbacks: { 
        label: item => `${item.label}: ${formatCurrency(item.raw)}` 
    } 
}
```

#### 🔄 Resize Automático

Ao trocar para a aba Dashboard:

```javascript
if (tabId === 'tab-dashboard') {
    Object.values(state.charts).forEach(c => c.resize());
}
```

---

## 📋 Módulo: Tabela Analítica

### 🏗️ Estrutura HTML

```html
<table class="siga-table">
  <thead>  ← Sticky, com cabeçalho azul escuro
    <tr>
      <th class="col-actions">🗑️</th>
      <th draggable="true">Coluna 1</th>
      ...
    </tr>
  </thead>
  <tbody>
    <tr>
      <td class="col-actions">🗑️</td>
      <td class="editable-cell">Valor</td>
      ...
    </tr>
  </tbody>
</table>
```

### 🎯 Recursos da Tabela

| Recurso | Como Ativar |
|---------|-------------|
| **Ordenação** | Clique no `<th>` (asc → desc → asc) |
| **Reordenação** | Arraste o `<th>` (drag & drop) |
| **Edição** | Duplo clique na célula |
| **Exclusão** | Clique no 🗑️ da linha |
| **Paginação** | Seletor `15 / 30 / 50 / 100 / Todos` |
| **Busca global** | Input no topo da tabela |
| **Colunas visíveis** | Dropdown suspenso |

### 📄 Paginação

**`renderPaginationButtons(totalPages, currentPage)`** gera botões:

```
[◀] [1] [2] [3] [4] [5] [▶]
     ↑ active (btn-primary)
```

Com janela deslizante: `Math.max(1, currentPage - 2)` até `Math.min(totalPages, currentPage + 2)`.

### 🎨 Classes CSS de Células

| Classe | Aplicação |
|--------|-----------|
| `.editable-cell` | Toda célula editável |
| `.num-cell` | Colunas com "VALOR" → `text-align: right` |
| `.date-cell` | Colunas de data → `text-align: center` |
| `.col-actions` | Coluna de exclusão |

### 🔤 Formatação de Exibição: `formatDisplayCell()`

Pipeline de formatação por tipo:

```javascript
function formatDisplayCell(headerName, rawVal, rowObj) {
    // 1. Vazio → ""
    if (rawVal === null || rawVal === undefined || String(rawVal).trim() === "") return "";
    
    // 2. VALOR → R$ com 2 decimais
