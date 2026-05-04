# ─────────────────────────────────────────────────────────────
#  Configuração da aplicação
#  Edite este ficheiro para pré-configurar a API key Anthropic
#  e outras definições globais.
# ─────────────────────────────────────────────────────────────

# API key Anthropic — pré-configurada, dispensando introdução manual na app.
# Se estiver vazia (""), a app mostrará um campo de texto para introdução manual.
ANTHROPIC_API_KEY = "sk-ant-..."   # ← coloque aqui a sua chave: "sk-ant-..."

# Modelo a usar para geração de texto
ANTHROPIC_MODEL = "claude-opus-4-5"

# Número máximo de tokens para cada chamada à API
MAX_TOKENS = 3200
