# 📞 Gerador de Relatório de Atendimento Telefónico

Aplicação web que transforma automaticamente um ficheiro Excel de dados de chamadas num relatório profissional e editável em Word (.docx), com análise por Inteligência Artificial.

---

## ✨ Funcionalidades

- **Leitura automática** do ficheiro Excel de exportação da centralita
- **4 gráficos** gerados automaticamente:
  - Distribuição global (atendidas vs. perdidas)
  - Top 15 grupos por volume
  - Top 10 grupos por Índice de Prioridade
  - Distribuição da taxa de resposta
- **Tabelas** detalhadas por grupo, ordenadas por Índice de Prioridade
- **Análise por IA** (introdução, interpretação, conclusões e recomendações)
- **Até 3 hipóteses** de relatório com estilos diferentes:
  - Hipótese 1 — Técnico e Formal
  - Hipótese 2 — Analítico-Estratégico
  - Hipótese 3 — Executivo e Conciso
- **Output Word editável** (.docx) com cabeçalho, rodapé e design minimalista

---

## ⚙️ Instalação

### 1. Pré-requisitos

- Python 3.9 ou superior
- Pip

### 2. Instalar dependências

```bash
pip install -r requirements.txt
```

### 3. Configurar a API Key da Anthropic

Opção A — variável de ambiente (recomendado):
```bash
export ANTHROPIC_API_KEY="sk-ant-..."
```

Opção B — introduzir diretamente na interface da aplicação.

---

## 🚀 Executar a aplicação

```bash
streamlit run app.py
```

A aplicação abre automaticamente no browser em `http://localhost:8501`.

---

## 📂 Formato do ficheiro de entrada

O ficheiro Excel deve ter exactamente **duas folhas**:

| Folha | Conteúdo |
|-------|----------|
| `Summary` | Totais globais do período (apresentadas, atendidas, perdidas, taxas) |
| `NO SUB GROUP` | Detalhe por grupo/extensão com todas as métricas |

O intervalo de datas é lido automaticamente da folha `Summary` (linha 3).

---

## 📋 Métricas calculadas

| Métrica | Fórmula |
|---------|---------|
| Taxa de Resposta | `Atendidas / Recebidas × 100` |
| Índice de Prioridade | `Chamadas Perdidas × (1 − Taxa de Resposta)` |

O **Índice de Prioridade** hierarquiza os grupos que exigem intervenção mais urgente: quanto mais alto, maior a combinação de volume de chamadas perdidas com baixa taxa de resposta.

---

## 📄 Estrutura do relatório gerado

1. Capa com KPIs em destaque
2. Introdução (gerada por IA)
3. Resumo Estatístico (tabela global)
4. Análise por Grupo (tabela detalhada ordenada por Índice de Prioridade)
5. Gráficos e Visualizações (4 gráficos)
6. Análise e Interpretação dos Dados (gerada por IA)
7. Conclusões e Recomendações (gerada por IA)

---

## 🗂️ Ficheiros

```
report_app/
├── app.py            # Aplicação principal
├── requirements.txt  # Dependências Python
└── README.md         # Este ficheiro
```

---

## 🗺️ Mapeamento de Grupos (grupos_mapping.json)

O ficheiro  converte os códigos internos  nos nomes reais dos assuntos/temas da centralita. Estrutura de cada entrada:

```json
"GeralM501": {
  "assunto":       "Benefício Fiscal ao Gasóleo Agrícola",
  "assunto_curto": "Benefício Fiscal ao Gasóleo Agrícola",
  "area":          "Ajuda de Assuntos",
  "servico":       "Linha de Apoio a Assuntos Agrícolas",
  "extensao":      4501
}
```

Para **atualizar o mapeamento** quando a estrutura IVR mudar, basta correr o script  com o novo ficheiro Excel de configuração IVR.
