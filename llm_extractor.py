"""
Extração de documentos cartorários com embeddings e JSON estruturado.

Um único arquivo com tudo: detecção, extração, embedding e export.
Extensível para novos tipos de documentos no futuro.
"""

import hashlib
import json
import logging
import os
import re
import time
import uuid

from dataclasses import dataclass
from datetime import datetime
from pathlib import Path
from typing import Dict, List, Optional

from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.shared import Pt
from dotenv import load_dotenv
from google import genai
from google.genai import types

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s - %(name)s - %(levelname)s - %(message)s",
    datefmt="%Y-%m-%d %H:%M:%S",
)
logger = logging.getLogger(__name__)
load_dotenv()

# =============================================================================
# PREÇOS (Gemini Developer API / AI Studio)
# =============================================================================
# gemini-2.5-flash
FLASH_INPUT_USD_PER_1M = 0.30
FLASH_OUTPUT_USD_PER_1M = 2.50

# Embedding (referência; dependendo da conta/produto, pode variar)
EMBED_INPUT_USD_PER_1M = 0.15

# Câmbio para estimativa (ajuste ou mova para .env se quiser)
USD_TO_BRL = float(os.getenv("USD_TO_BRL", "5.25"))

# =============================================================================
# Prompt para extração
# =============================================================================

EXTRACTION_PROMPT = """Você receberá uma imagem de um documento de registro de imóveis de cartório brasileiro.
INSTRUÇÕES CRÍTICAS:

1. EXTRAIA O TEXTO COM MÁXIMA FIDELIDADE:
   - Reproduza EXATAMENTE como aparece no documento
   - Preserva estrutura, quebras de linha, espaçamento
   - Mantenha números, datas e valores exatamente como estão

2. PALAVRAS/TRECHOS ILEGÍVEIS OU COM BAIXA CONFIANÇA:
   - Se não tem CERTEZA ABSOLUTA, substitua por ***
   - NUNCA repita palavras ou trechos por dúvida
   - Se vê repetição suspeita, é provável erro de OCR - substitua por ***

3. PADRÕES REPETITIVOS SUSPEITOS:
   - Se a mesma frase/número se repete muitas vezes (>2x), é erro do OCR
   - Mantenha APENAS a primeira ocorrência legítima
   - Substitua repetições suspeitas por ***

4. QUALIDADE DO DOCUMENTO - SEJA EXTREMAMENTE CONSERVADOR:
   - Se a qualidade geral é muito ruim (muitas letras/palavras ilegíveis):
     * Marque seções inteiras como *** quando não conseguir ler com confiança
     * NUNCA tente "adivinhar" o que está escrito
     * Melhor marcar como *** do que retornar lixo/texto errado
   - Para cada parágrafo: se >30% das palavras estão ilegíveis, considere substituir TUDO por ***

5. PÓS-PROCESSAMENTO DE ORTOGRAFIA (APENAS COM MÁXIMA CONFIANÇA):
   - APENAS corrija erros óbvios de OCR em palavras muito comuns
   - Exemplos de correção permitida:
     * "rna" → "uma" (OCR comum)
     * "senhor" → "senhor" (já correto, não mude)
     * "l" → "I" em nomes próprios (não mude, pode ser intenção)
   - NUNCA corrija:
     * Nomes próprios (mesmo que pareça erro)
     * Palavras que não tem 100% de certeza
     * Palavras arcaicas ou regionais
     * Se tiver dúvida, deixe como está ou substitua por ***

6. IMPORTANTE - NÃO FAZER:
   - Não "complete" palavras incompletas
   - Não corrija abreviações
   - Não reorganize o texto
   - Não invente conteúdo por ter baixa confiança
   - NÃO repita o que não tem certeza
   - NÃO corrija ortografia por dúvida (risco maior que erro mantido)
   - SE NÃO TEM CERTEZA, PREFIRA *** - ISSO É MELHOR QUE RETORNAR TEXTO ERRADO

Retorne APENAS o texto extraído, sem avisos ou comentários."""

# =============================================================================
# CONFIGURAÇÃO DE CAMPOS CARTORÁRIOS
# =============================================================================


@dataclass
class CampoExtracao:
    """Definição de um campo a extrair."""

    chave: str
    regex: str
    descricao: str


class DocumentConfigurations:
    """Configurações de extração para documentos cartorários."""

    CAMPOS_COMUNS = [
        CampoExtracao("data_documento", r"(\d{1,2})[/-](\d{1,2})[/-](\d{2,4})", "Data"),
        CampoExtracao("cpf", r"(\d{3})[.\-](\d{3})[.\-](\d{3})[.\-](\d{2})", "CPF"),
    ]

    REGISTRO_IMOVEIS_CAMPOS = [
        CampoExtracao("matricula", r"MATRÍCULA\s+N\.?\s*[:\s]*(\d+)", "Matrícula"),
        CampoExtracao("livro", r"LIVRO\s+N\.?\s*[:\s]*([0-9A-Z\s]+)", "Livro"),
        CampoExtracao("folha", r"FL\.?\s+N\.?\s*[:\s]*(\d+)", "Folha"),
        CampoExtracao(
            "proprietario",
            r"(?:Proprietário|Nome)[:\s]+([A-Za-záéíóú\s]+?)(?:\n|,|$)",
            "Proprietário",
        ),
    ]

    @classmethod
    def get_campos_por_tipo(cls, tipo: str) -> List[CampoExtracao]:
        """Retorna campos para extração baseado no tipo de documento."""
        if tipo == "Registro de Imóveis":
            return cls.CAMPOS_COMUNS + cls.REGISTRO_IMOVEIS_CAMPOS
        # Adicione novos tipos aqui quando necessário
        return cls.CAMPOS_COMUNS


# =============================================================================
# DETECÇÃO DE TIPO
# =============================================================================


class DocumentTypeDetector:
    """Detecta tipo de documento com padrões configuráveis."""

    PADROES = {
        "Registro de Imóveis": ["REGISTRO DE IMÓVEIS", "LIVRO N", "MATRÍCULA N"],
        # Novos tipos aqui: "Título": [...], "Procuração": [...]
    }

    @staticmethod
    def detectar_tipo(texto: str) -> Dict:
        """Detecta tipo de documento analisando padrões."""
        texto_upper = texto.upper()

        for tipo, padroes in DocumentTypeDetector.PADROES.items():
            matches = sum(1 for p in padroes if p in texto_upper)
            if matches > 0:
                confianca = min(0.95 + matches * 0.01, 1.0)
                logger.info("%s detectado (confiança: %.2f)", tipo, confianca)
                return {
                    "tipo": tipo,
                    "confianca": confianca,
                    "categorias": ["cartório", "imóvel"],
                }

        logger.warning("Tipo não reconhecido, assumindo Registro de Imóveis")
        return {
            "tipo": "Registro de Imóveis",
            "confianca": 0.5,
            "categorias": ["cartório", "imóvel"],
        }


# =============================================================================
# EXTRAÇÃO DE CAMPOS
# =============================================================================


class FieldExtractor:
    """Extrator de campos usando regex."""

    @staticmethod
    def extrair(texto: str, tipo: str) -> Dict[str, str]:
        """Extrai campos estruturados do texto."""
        campos = DocumentConfigurations.get_campos_por_tipo(tipo)
        resultado: Dict[str, str] = {}

        for campo in campos:
            try:
                match = re.search(campo.regex, texto, re.IGNORECASE)
                if match:
                    valor = (
                        match.group(1).strip()
                        if match.lastindex
                        else match.group(0).strip()
                    )
                    resultado[campo.chave] = valor
            except re.error:
                logger.warning("Erro no regex para %s", campo.chave)

        if not resultado:
            resultado["observacao"] = "Nenhum campo específico extraído"

        logger.info("Campos extraídos: %d", len(resultado))
        return resultado


# =============================================================================
# UTILITÁRIOS
# =============================================================================


def get_mime_type(file_path: str) -> str:
    """Detecta tipo MIME da imagem."""
    ext = Path(file_path).suffix.lower()
    mime_types = {
        ".jpg": "image/jpeg",
        ".jpeg": "image/jpeg",
        ".png": "image/png",
        ".webp": "image/webp",
        ".tiff": "image/tiff",
        ".tif": "image/tiff",
    }
    return mime_types.get(ext, "image/jpeg")


def analyze_legibility(text: str) -> Dict:
    """Analisa legibilidade do texto extraído."""
    illegible = text.count("***")
    words = text.split()
    total = len(words)

    if total == 0:
        return {
            "illegible_count": 0,
            "total_words": 0,
            "legibility_percentage": 0,
            "has_warning": True,
        }

    legible = total - (illegible * 3)
    percentage = (legible / total) * 100 if total > 0 else 0

    return {
        "illegible_count": illegible,
        "total_words": total,
        "legibility_percentage": round(percentage, 2),
        "has_warning": percentage < 50,
    }


def calculate_hash(file_path: str) -> str:
    """Calcula hash SHA256 do arquivo."""
    sha256 = hashlib.sha256()
    with open(file_path, "rb") as f:
        for chunk in iter(lambda: f.read(4096), b""):
            sha256.update(chunk)
    return sha256.hexdigest()


def extract_usage_metadata(resp) -> dict:
    """
    Extrai contagem de tokens da resposta do SDK do Google.
    O formato pode variar por versão/modelo, então tentamos múltiplas rotas.
    Retorna dict possivelmente com:
      - prompt_token_count
      - candidates_token_count
      - total_token_count
      - cached_content_token_count
    """
    usage: dict = {}

    um = getattr(resp, "usage_metadata", None)
    if um is not None:
        # Caso seja objeto com atributos
        for k in (
            "prompt_token_count",
            "candidates_token_count",
            "total_token_count",
            "cached_content_token_count",
        ):
            if hasattr(um, k):
                usage[k] = getattr(um, k)
        # Caso seja dict
        if not usage and isinstance(um, dict):
            usage = dict(um)

    # Fallback: tenta model_dump (pydantic) se existir
    if not usage:
        try:
            d = resp.model_dump()
        except Exception:
            d = None
        if isinstance(d, dict):
            um2 = d.get("usage_metadata") or d.get("usageMetadata") or d.get("usage")
            if isinstance(um2, dict):
                usage = dict(um2)

    return usage


def estimate_flash_cost_usd(usage: dict) -> Optional[dict]:
    """
    Estima custo do gemini-2.5-flash baseado em usage_metadata.
    Como total_token_count inclui imagem+prompt+etc, usamos:
      output = candidates_token_count
      input  = total_token_count - candidates_token_count
    """
    if not usage:
        return None

    total = usage.get("total_token_count")
    out_toks = usage.get("candidates_token_count")

    if not isinstance(total, int) or not isinstance(out_toks, int):
        return None

    in_toks = max(0, total - out_toks)

    cost_usd = (in_toks / 1_000_000) * FLASH_INPUT_USD_PER_1M + (
        out_toks / 1_000_000
    ) * FLASH_OUTPUT_USD_PER_1M
    return {
        "input_tokens_est": in_toks,
        "output_tokens": out_toks,
        "total_tokens": total,
        "cost_usd": cost_usd,
        "cost_brl": cost_usd * USD_TO_BRL,
    }


def estimate_embed_cost_usd(text_token_estimate: Optional[int]) -> Optional[dict]:
    """
    Estima custo de embedding por tokens de entrada (se você estimar/medir tokens).
    Como o embed nem sempre retorna usage_metadata, deixamos opcional.
    """
    if not isinstance(text_token_estimate, int) or text_token_estimate <= 0:
        return None
    cost_usd = (text_token_estimate / 1_000_000) * EMBED_INPUT_USD_PER_1M
    return {
        "input_tokens_est": text_token_estimate,
        "cost_usd": cost_usd,
        "cost_brl": cost_usd * USD_TO_BRL,
    }


def generate_embedding(text: str, api_key: str) -> Optional[list]:
    """Gera embedding do texto via Gemini API (gemini-embedding-001)."""
    try:
        client = genai.Client(api_key=api_key)

        # Gemini Developer API embedding model
        result = client.models.embed_content(
            model="gemini-embedding-001",
            contents=text[:8000],  # ajuste depois se quiser
        )

        # Alguns SDKs retornam 'embeddings' e outros 'embedding'
        emb = None
        if hasattr(result, "embeddings") and result.embeddings:
            emb = result.embeddings[0]
        elif hasattr(result, "embedding"):
            emb = result.embedding

        if emb is not None and hasattr(emb, "values"):
            logger.info("Embedding gerado (%d dimensões)", len(emb.values))
            return emb.values

        logger.warning("Embedding retornou vazio/inesperado")
        return None

    except Exception as e:
        logger.warning("Erro ao gerar embedding: %s", e)
        return None


# =============================================================================
# EXTRAÇÃO PRINCIPAL
# =============================================================================


def extract_from_document(
    image_path: str, api_key: str, model: str = "gemini-2.5-flash"
) -> Dict:
    """Extrai texto de documento usando Gemini API."""
    start_time = time.time()
    usage: dict = {}

    try:
        if not os.path.exists(image_path):
            logger.error("Arquivo não encontrado: %s", image_path)
            return {
                "success": False,
                "error": f"Arquivo não encontrado: {image_path}",
                "usage": {},
            }

        logger.info("Processando: %s", image_path)

        client = genai.Client(api_key=api_key)
        with open(image_path, "rb") as f:
            image_data = f.read()

        mime = get_mime_type(image_path)
        image_part = types.Part.from_bytes(data=image_data, mime_type=mime)

        api_start = time.time()
        response = client.models.generate_content(
            model=model, contents=[EXTRACTION_PROMPT, image_part]
        )
        api_time = time.time() - api_start

        usage = extract_usage_metadata(response)
        if usage:
            logger.info("TOKENS (%s): %s", model, usage)
        else:
            logger.warning(
                "TOKENS: usage_metadata não veio na resposta (pode variar por SDK/modelo)."
            )

        extracted = response.text or ""
        file_name = Path(image_path).name
        legibility = analyze_legibility(extracted)
        file_hash = calculate_hash(image_path)
        classification = DocumentTypeDetector.detectar_tipo(extracted)
        campos = FieldExtractor.extrair(extracted, classification["tipo"])
        embedding = generate_embedding(extracted, api_key)
        processing_time = time.time() - start_time

        if legibility["has_warning"]:
            logger.warning(
                "Legibilidade baixa: %.2f%%", legibility["legibility_percentage"]
            )
        else:
            logger.info("Legibilidade: %.2f%%", legibility["legibility_percentage"])

        logger.info(
            "Tipo: %s | API: %.2fs | Total: %.2fs",
            classification["tipo"],
            api_time,
            processing_time,
        )

        return {
            "success": True,
            "id": str(uuid.uuid4()),
            "arquivo_original": file_name,
            "arquivo_hash": file_hash,
            "criado_em": datetime.now().isoformat(),
            "extracted_text": extracted,
            "file_name": file_name,
            "model_used": model,
            "legibility_analysis": legibility,
            "processing_time": processing_time,
            "api_time": api_time,
            "classificacao": classification,
            "campos_extraidos": campos,
            "embedding": embedding,
            "usage": usage,  # <-- CRÍTICO: agora o main consegue somar e calcular custo
        }

    except Exception as e:
        processing_time = time.time() - start_time
        logger.error("Erro ao processar documento: %s", e)
        return {
            "success": False,
            "error": str(e),
            "file_name": Path(image_path).name
            if os.path.exists(image_path)
            else image_path,
            "processing_time": processing_time,
            "usage": usage or {},
        }


# =============================================================================
# EXPORTAÇÃO
# =============================================================================


def export_to_word(result: Dict, output_path: Optional[str] = None) -> str:
    """Exporta para Word (.docx)."""
    if not result["success"]:
        raise ValueError("Não é possível exportar resultado falho")

    if output_path is None:
        stem = Path(result["file_name"]).stem
        output_path = f"saida/extracted_{stem}.docx"

    Path(output_path).parent.mkdir(parents=True, exist_ok=True)

    doc = Document()
    title = doc.add_heading("Documento Extraído", level=1)
    title.alignment = WD_ALIGN_PARAGRAPH.CENTER
    doc.add_paragraph()

    for line in result["extracted_text"].split("\n"):
        line = line.strip()
        if not line:
            continue

        is_section = (line.isupper() and len(line) > 3) or "Nº" in line or "Fl." in line

        if is_section and not line.startswith("***"):
            para = doc.add_heading(line, level=3)
        else:
            para = doc.add_paragraph(line)

        if hasattr(para, "paragraph_format"):
            para.paragraph_format.line_spacing = 1.15
            para.paragraph_format.space_after = Pt(6)

        for run in para.runs:
            run.font.size = Pt(11)

    doc.save(output_path)
    logger.info("Word salvo: %s", output_path)
    return output_path


def export_to_json(result: Dict, output_path: Optional[str] = None) -> str:
    """Exporta para JSON com metadados e embedding."""
    if not result["success"]:
        raise ValueError("Não é possível exportar resultado falho")

    if output_path is None:
        stem = Path(result["file_name"]).stem
        output_path = f"saida/extracted_{stem}.json"

    Path(output_path).parent.mkdir(parents=True, exist_ok=True)

    usage = result.get("usage") or {}

    json_data = {
        "id": result["id"],
        "arquivo_original": result["arquivo_original"],
        "arquivo_hash": result["arquivo_hash"],
        "criado_em": result["criado_em"],
        "modelo_utilizado": result["model_used"],
        "processamento": {
            "tempo_total_segundos": result["processing_time"],
            "tempo_api_segundos": result["api_time"],
        },
        "tokens": usage,
        "legibilidade": {
            "palavras_ilegíveis": result["legibility_analysis"]["illegible_count"],
            "total_palavras": result["legibility_analysis"]["total_words"],
            "percentual_legibilidade": result["legibility_analysis"][
                "legibility_percentage"
            ],
            "alerta": result["legibility_analysis"]["has_warning"],
        },
        "classificacao": {
            "tipo_documento": result["classificacao"]["tipo"],
            "confianca": result["classificacao"]["confianca"],
            "categorias": result["classificacao"]["categorias"],
        },
        "campos_extraidos": result["campos_extraidos"],
        "embedding": result["embedding"],
        "texto_extraido": result["extracted_text"],
    }

    with open(output_path, "w", encoding="utf-8") as f:
        json.dump(json_data, f, ensure_ascii=False, indent=2)

    logger.info("JSON salvo: %s", output_path)
    return output_path


# =============================================================================
# MAIN
# =============================================================================


def main():
    """Executa pipeline completo (lote de 5 imagens) e reporta tokens e custo."""
    main_start = time.time()
    logger.info("Iniciando em %s", datetime.now().strftime("%Y-%m-%d %H:%M:%S"))

    api_key = os.getenv("GEMINI_API_KEY")
    if not api_key:
        logger.error("GEMINI_API_KEY não configurada no .env")
        return

    image_paths = [
        "img/1-img.jpeg",
        "img/2-img.jpeg",
        "img/3-img.jpeg",
        "img/4-img.jpeg",
        "img/5-img.jpeg",
    ]

    totals = {
        "prompt_token_count": 0,
        "candidates_token_count": 0,
        "total_token_count": 0,
        "cached_content_token_count": 0,
    }

    total_cost_usd = 0.0
    total_cost_brl = 0.0

    ok = 0
    for document_path in image_paths:
        if not os.path.exists(document_path):
            logger.error("Arquivo não encontrado: %s", document_path)
            continue

        result = extract_from_document(document_path, api_key)

        # fallback de modelo (se estourar quota/limite)
        if not result["success"] and "RESOURCE_EXHAUSTED" in str(
            result.get("error", "")
        ):
            logger.warning("Tentando modelo alternativo...")
            result = extract_from_document(
                document_path, api_key, model="gemini-1.5-flash-8b"
            )

        if result["success"]:
            ok += 1
            usage = result.get("usage") or {}

            # soma tokens do lote (quando vierem)
            for k in totals.keys():
                v = usage.get(k)
                if isinstance(v, int):
                    totals[k] += v

            # custo estimado do Flash (quando der para estimar)
            cost = estimate_flash_cost_usd(usage)
            if cost:
                total_cost_usd += cost["cost_usd"]
                total_cost_brl += cost["cost_brl"]
                logger.info(
                    "CUSTO ESTIMADO: %s | in=%d out=%d total=%d | US$ %.6f | R$ %.4f",
                    document_path,
                    cost["input_tokens_est"],
                    cost["output_tokens"],
                    cost["total_tokens"],
                    cost["cost_usd"],
                    cost["cost_brl"],
                )
            else:
                logger.warning(
                    "CUSTO ESTIMADO: sem usage suficiente para calcular (%s)",
                    document_path,
                )

            # se você quiser exportar para cada imagem, descomente:
            # export_to_word(result)
            # export_to_json(result)

            logger.info(
                "OK: %s | api_time=%.2fs | total_tokens=%s",
                document_path,
                result.get("api_time", -1),
                usage.get("total_token_count"),
            )
        else:
            logger.error("Falha em %s: %s", document_path, result.get("error"))

    total_time = time.time() - main_start
    logger.info("========================================")
    logger.info("Resumo lote: %d/%d ok | %.2fs", ok, len(image_paths), total_time)
    logger.info("TOKENS TOTAIS (lote): %s", totals)

    if ok > 0:
        logger.info(
            "CUSTO TOTAL (LLM Flash): US$ %.6f | R$ %.4f | média por doc: US$ %.6f | R$ %.4f",
            total_cost_usd,
            total_cost_brl,
            total_cost_usd / ok,
            total_cost_brl / ok,
        )
    else:
        logger.info("CUSTO TOTAL (LLM Flash): sem documentos processados com sucesso.")


if __name__ == "__main__":
    main()
