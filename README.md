# OCR Cartório - Extração de Documentos com LLM

Um sistema completo para extração, processamento e análise de documentos cartorários brasileiros usando a API Gemini com embeddings e estruturação de dados em JSON.

---

## 📋 Visão Geral

O arquivo `llm_extractor.py` contém um pipeline integrado que:

1. **Extrai** texto de imagens de documentos usando a API Gemini
2. **Detecta** o tipo de documento automaticamente
3. **Extrai campos** estruturados usando regex
4. **Analisa** a legibilidade do texto extraído
5. **Gera** embeddings para busca semântica
6. **Exporta** resultados em Word e JSON

---

## 🔧 Estrutura de Classes e Funções

### **Classe: `CampoExtracao`** (dataclass)
Define a estrutura de um campo a ser extraído.

**Atributos:**
- `chave`: identificador único do campo (ex: "cpf", "data_documento")
- `regex`: padrão de regex para buscar o campo no texto
- `descricao`: descrição legível do campo

**Exemplo de uso:**
```python
campo = CampoExtracao("cpf", r"(\d{3})[.\-](\d{3})[.\-](\d{3})[.\-](\d{2})", "CPF")
```

---

### **Classe: `DocumentConfigurations`**
Centralizador de configurações de campos por tipo de documento.

**Atributos de classe:**
- `CAMPOS_COMUNS`: campos presentes em todos os documentos (data, CPF)
- `REGISTRO_IMOVEIS_CAMPOS`: campos específicos de registro de imóveis

**Método:**
- `get_campos_por_tipo(tipo: str) → List[CampoExtracao]`
  - **Descrição:** retorna a lista de campos a extrair baseado no tipo de documento
  - **Parâmetro:** tipo do documento (ex: "Registro de Imóveis")
  - **Retorno:** lista de campos configurados para aquele tipo

---

### **Classe: `DocumentTypeDetector`**
Detecta automaticamente que tipo de documento está sendo processado.

**Atributo de classe:**
- `PADROES`: dicionário com padrões de texto para cada tipo de documento

**Método:**
- `detectar_tipo(texto: str) → Dict`
  - **Descrição:** analisa o texto extraído procurando por padrões característicos
  - **Parâmetro:** texto extraído do documento
  - **Retorno:** dicionário com `tipo`, `confianca` (0-1) e `categorias`
  - **Exemplo:**
    ```python
    resultado = DocumentTypeDetector.detectar_tipo("REGISTRO DE IMÓVEIS LIVRO N 100")
    # Retorna: {"tipo": "Registro de Imóveis", "confianca": 0.98, "categorias": [...]}
    ```

---

### **Classe: `FieldExtractor`**
Extrai campos estruturados do texto usando padrões regex.

**Método:**
- `extrair(texto: str, tipo: str) → Dict[str, str]`
  - **Descrição:** busca campos específicos no texto baseado no tipo de documento
  - **Parâmetros:**
    - `texto`: texto extraído do documento
    - `tipo`: tipo de documento detectado
  - **Retorno:** dicionário com chave → valor dos campos encontrados
  - **Exemplo:**
    ```python
    campos = FieldExtractor.extrair(texto, "Registro de Imóveis")
    # Retorna: {"matricula": "1234", "cpf": "123.456.789-00", ...}
    ```

---

## 🛠️ Funções Utilitárias

### `get_mime_type(file_path: str) → str`
**Descrição:** detecta o tipo MIME (media type) da imagem baseado na extensão.

**Parâmetro:** caminho do arquivo de imagem

**Retorno:** string com tipo MIME (ex: "image/jpeg", "image/png")

**Exemplo:**
```python
mime = get_mime_type("documento.jpg")  # Retorna: "image/jpeg"
```

---

### `analyze_legibility(text: str) → Dict`
**Descrição:** analisa a qualidade de legibilidade do texto extraído.

Conta ocorrências de `***` (indicadores de texto ilegível) e calcula a percentagem de legibilidade.

**Parâmetro:** texto extraído do documento

**Retorno:** dicionário com:
- `illegible_count`: número de blocos ilegíveis (`***`)
- `total_words`: quantidade total de palavras
- `legibility_percentage`: percentual de legibilidade (0-100)
- `has_warning`: booleano indicando se a legibilidade está abaixo de 50%

**Exemplo:**
```python
legibility = analyze_legibility("Este é um texto com *** ilegível")
# Retorna: {
#   "illegible_count": 1,
#   "total_words": 7,
#   "legibility_percentage": 85.71,
#   "has_warning": False
# }
```

---

### `calculate_hash(file_path: str) → str`
**Descrição:** calcula o hash SHA256 do arquivo original para verificação de integridade.

**Parâmetro:** caminho do arquivo de imagem

**Retorno:** string com hash SHA256 em formato hexadecimal

**Exemplo:**
```python
hash_valor = calculate_hash("img/documento.jpg")
# Retorna: "a7f3b2c1d4e5f6g7h8i9j0k1l2m3n4o5p6q7r8s9t0u1v2w3x4y5z..."
```

---

### `generate_embedding(text: str, api_key: str) → Optional[list]`
**Descrição:** gera um embedding (representação vetorial) do texto usando a API Gemini.

Retorna um vetor de 1536 dimensões que pode ser usado para buscas semânticas.

**Parâmetros:**
- `text`: texto a ser convertido em embedding
- `api_key`: chave de API do Google Gemini

**Retorno:** lista com 1536 valores float (embedding) ou None se houver erro

**Nota:** o texto é truncado em 2000 caracteres antes do processamento

---

### `extract_from_document(image_path: str, api_key: str, model: str = "gemini-2.5-flash") → Dict`
**Descrição:** função principal que executa todo o pipeline de extração em um documento.

**Parâmetros:**
- `image_path`: caminho da imagem do documento
- `api_key`: chave de API do Google Gemini
- `model`: modelo Gemini a usar (padrão: "gemini-2.5-flash")

**Retorno:** dicionário abrangente com:
- `success`: booleano indicando sucesso
- `id`: UUID único do processamento
- `arquivo_original`: nome do arquivo
- `arquivo_hash`: hash SHA256 do arquivo
- `criado_em`: timestamp ISO 8601
- `extracted_text`: texto completo extraído
- `model_used`: modelo Gemini utilizado
- `legibility_analysis`: análise de legibilidade
- `processing_time`: tempo total em segundos
- `api_time`: tempo de resposta da API em segundos
- `classificacao`: detecção de tipo e confiança
- `campos_extraidos`: dicionário com campos estruturados
- `embedding`: vetor de 1536 dimensões

**Exemplo:**
```python
resultado = extract_from_document("img/documento.jpg", api_key)
if resultado["success"]:
    print(f"Tipo: {resultado['classificacao']['tipo']}")
    print(f"Legibilidade: {resultado['legibility_analysis']['legibility_percentage']}%")
```

---

### `export_to_word(result: Dict, output_path: Optional[str] = None) → str`
**Descrição:** exporta o texto extraído para um arquivo Word (.docx) formatado.

O documento Word é formatado com:
- Título centralizado
- Seções detectadas automaticamente como headings
- Espaçamento e fonte adequados
- Linhas quebradas preservadas

**Parâmetros:**
- `result`: dicionário retornado por `extract_from_document()`
- `output_path`: caminho de saída (opcional, padrão: `saida/extracted_{nome}.docx`)

**Retorno:** caminho do arquivo Word criado

**Lança exceção:** se o resultado indicar falha na extração

**Exemplo:**
```python
caminho_word = export_to_word(resultado)
# Cria: saida/extracted_documento.docx
```

---

### `export_to_json(result: Dict, output_path: Optional[str] = None) → str`
**Descrição:** exporta os dados completos da extração para JSON estruturado.

O JSON inclui:
- Metadados (ID, hash, timestamp)
- Análise de legibilidade
- Classificação do documento
- Campos estruturados extraídos
- Embedding (vetor 1536-d)
- Texto completo extraído

**Parâmetros:**
- `result`: dicionário retornado por `extract_from_document()`
- `output_path`: caminho de saída (opcional, padrão: `saida/extracted_{nome}.json`)

**Retorno:** caminho do arquivo JSON criado

**Lança exceção:** se o resultado indicar falha na extração

**Exemplo:**
```python
caminho_json = export_to_json(resultado)
# Cria: saida/extracted_documento.json
```

---

### `main()`
**Descrição:** função principal que orquestra o pipeline completo.

**Fluxo:**
1. Lê a chave API do Gemini do arquivo `.env`
2. Carrega o documento em `img/3-img.jpeg`
3. Executa `extract_from_document()`
4. Se falhar com erro de limite de taxa, tenta modelo alternativo
5. Se suceder, exporta para Word e JSON
6. Registra logs com tempos de processamento

**Retorno:** None (resultados salvos em arquivos)

---

## 📊 Estrutura do JSON Exportado

```json
{
  "id": "uuid-aqui",
  "arquivo_original": "3-img.jpeg",
  "arquivo_hash": "sha256-aqui",
  "criado_em": "2025-01-31T10:30:45.123456",
  "modelo_utilizado": "gemini-2.5-flash",
  "processamento": {
    "tempo_total_segundos": 5.234,
    "tempo_api_segundos": 4.891
  },
  "legibilidade": {
    "palavras_ilegíveis": 2,
    "total_palavras": 156,
    "percentual_legibilidade": 98.72,
    "alerta": false
  },
  "classificacao": {
    "tipo_documento": "Registro de Imóveis",
    "confianca": 0.98,
    "categorias": ["cartório", "imóvel"]
  },
  "campos_extraidos": {
    "matricula": "1234",
    "cpf": "123.456.789-00",
    "data_documento": "15/01/2025"
  },
  "embedding": [0.123, 0.456, ...],
  "texto_extraido": "REGISTRO DE IMÓVEIS..."
}
```

---

## 🚀 Como Usar

### Setup Inicial

```bash
# Instalar dependências
pip install -r req.txt

# Criar arquivo .env com a chave da API
echo "GEMINI_API_KEY=sua_chave_aqui" > .env
```

### Processar um Documento

```python
from llm_extractor import extract_from_document, export_to_word, export_to_json
import os

api_key = os.getenv("GEMINI_API_KEY")

# Extrair
resultado = extract_from_document("img/seu_documento.jpg", api_key)

# Exportar
if resultado["success"]:
    export_to_word(resultado)
    export_to_json(resultado)
    print("✓ Documento processado com sucesso!")
```

---

## ⚙️ Configuração de Novos Tipos de Documentos

Para adicionar um novo tipo de documento:

1. Adicione padrões em `DocumentTypeDetector.PADROES`:
```python
PADROES = {
    "Registro de Imóveis": [...],
    "Título": ["TÍTULO N", "VALOR", "DESCRITIVO"],  # Novo
}
```

2. Adicione campos em `DocumentConfigurations`:
```python
TITULO_CAMPOS = [
    CampoExtracao("numero_titulo", r"TÍTULO\s+N\.?\s*[:\s]*(\d+)", "Número"),
    # ... mais campos
]
```

3. Atualize o método `get_campos_por_tipo()`:
```python
if tipo == "Título":
    return cls.CAMPOS_COMUNS + cls.TITULO_CAMPOS
```

---

## 📝 Variáveis de Ambiente

Crie um arquivo `.env` com:

```
GEMINI_API_KEY=sua_chave_de_api_aqui
```

---

## 📦 Dependências Principais

- `google-genai`: API do Google Gemini
- `python-docx`: criação de documentos Word
- `python-dotenv`: carregamento de variáveis de ambiente

---

## 📍 Saídas Geradas

- **Word (.docx)**: em `saida/extracted_{nome_arquivo}.docx`
- **JSON (.json)**: em `saida/extracted_{nome_arquivo}.json`

---

## 📊 Logs

O sistema registra informações detalhadas incluindo:
- Tipos de documentos detectados
- Tempos de processamento
- Análise de legibilidade
- Avisos sobre qualidade
- Erros e exceções

Logs aparecem no console com formato: `[timestamp] [nivel] mensagem`

