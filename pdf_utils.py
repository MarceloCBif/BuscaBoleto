"""
Utilitário para extração de dados de boletos PDF usando pdfplumber e regex.
"""

import re
from typing import Optional, Dict
import pdfplumber


class BoletoExtractor:
    """Extrai informações de boletos em PDF."""
    
    # Padrões regex para identificar dados do boleto
    PATTERNS = {
        # Linha digitável (47 dígitos com espaços)
        'linha_digitavel': r'\d{5}[\.\s]?\d{5}[\.\s]?\d{5}[\.\s]?\d{6}[\.\s]?\d{5}[\.\s]?\d{6}[\.\s]?\d{1}[\.\s]?\d{14}',
        
        # Código de barras (44 dígitos)
        'codigo_barras': r'\d{44}',
        
        # Valor do boleto (R$ X.XXX,XX)
        'valor': r'R\$\s*[\d\.,]+',
        
        # Data de vencimento (DD/MM/AAAA)
        'vencimento': r'\d{2}/\d{2}/\d{4}',
        
        # CNPJ (XX.XXX.XXX/XXXX-XX)
        'cnpj': r'\d{2}\.\d{3}\.\d{3}/\d{4}-\d{2}',
        
        # CPF (XXX.XXX.XXX-XX)
        'cpf': r'\d{3}\.\d{3}\.\d{3}-\d{2}',
        
        # Nosso número (varia por banco)
        'nosso_numero': r'[Nn]osso\s*[Nn][uú]mero[:\s]*(\d+[\d\.\-/]*\d*)',
        
        # Número do documento
        'numero_documento': r'[Nn][uú]mero\s*[Dd]o\s*[Dd]ocumento[:\s]*(\d+)',
    }
    
    def __init__(self, pdf_path: str):
        """
        Inicializa o extrator com o caminho do PDF.
        
        Args:
            pdf_path: Caminho para o arquivo PDF do boleto.
        """
        self.pdf_path = pdf_path
        self.texto_completo = ""
        self.dados_extraidos: Dict[str, Optional[str]] = {}
    
    def extrair_texto(self) -> str:
        """
        Extrai todo o texto do PDF.
        
        Returns:
            Texto completo do PDF.
        """
        try:
            with pdfplumber.open(self.pdf_path) as pdf:
                textos = []
                for pagina in pdf.pages:
                    texto = pagina.extract_text()
                    if texto:
                        textos.append(texto)
                
                self.texto_completo = "\n".join(textos)
                return self.texto_completo
        except Exception as e:
            print(f"Erro ao extrair texto do PDF: {e}")
            return ""
    
    def extrair_linha_digitavel(self) -> Optional[str]:
        """
        Extrai a linha digitável do boleto.
        
        Returns:
            Linha digitável encontrada ou None.
        """
        if not self.texto_completo:
            self.extrair_texto()
        
        # Procura padrão de linha digitável
        match = re.search(self.PATTERNS['linha_digitavel'], self.texto_completo)
        if match:
            # Remove espaços e pontos
            linha = re.sub(r'[\.\s]', '', match.group())
            self.dados_extraidos['linha_digitavel'] = linha
            return linha
        
        return None
    
    def extrair_valor(self) -> Optional[str]:
        """
        Extrai o valor do boleto.
        
        Returns:
            Valor encontrado ou None.
        """
        if not self.texto_completo:
            self.extrair_texto()
        
        matches = re.findall(self.PATTERNS['valor'], self.texto_completo)
        if matches:
            # Geralmente o último valor é o total do boleto
            valor = matches[-1] if len(matches) > 1 else matches[0]
            self.dados_extraidos['valor'] = valor
            return valor
        
        return None
    
    def extrair_vencimento(self) -> Optional[str]:
        """
        Extrai a data de vencimento do boleto.
        
        Returns:
            Data de vencimento encontrada ou None.
        """
        if not self.texto_completo:
            self.extrair_texto()
        
        matches = re.findall(self.PATTERNS['vencimento'], self.texto_completo)
        if matches:
            # A primeira data geralmente é o vencimento
            vencimento = matches[0]
            self.dados_extraidos['vencimento'] = vencimento
            return vencimento
        
        return None
    
    def extrair_nosso_numero(self) -> Optional[str]:
        """
        Extrai o nosso número do boleto.
        
        Returns:
            Nosso número encontrado ou None.
        """
        if not self.texto_completo:
            self.extrair_texto()
        
        match = re.search(self.PATTERNS['nosso_numero'], self.texto_completo, re.IGNORECASE)
        if match:
            nosso_numero = match.group(1).strip()
            self.dados_extraidos['nosso_numero'] = nosso_numero
            return nosso_numero
        
        return None
    
    def extrair_cnpj(self) -> Optional[str]:
        """
        Extrai o CNPJ do beneficiário.
        
        Returns:
            CNPJ encontrado ou None.
        """
        if not self.texto_completo:
            self.extrair_texto()
        
        match = re.search(self.PATTERNS['cnpj'], self.texto_completo)
        if match:
            cnpj = match.group()
            self.dados_extraidos['cnpj'] = cnpj
            return cnpj
        
        return None
    
    def extrair_todos_dados(self) -> Dict[str, Optional[str]]:
        """
        Extrai todos os dados disponíveis do boleto.
        
        Returns:
            Dicionário com todos os dados extraídos.
        """
        if not self.texto_completo:
            self.extrair_texto()
        
        self.extrair_linha_digitavel()
        self.extrair_valor()
        self.extrair_vencimento()
        self.extrair_nosso_numero()
        self.extrair_cnpj()
        
        return self.dados_extraidos
    
    def buscar_numero_no_texto(self, numero: str) -> bool:
        """
        Verifica se um número específico existe no texto do boleto.
        
        Args:
            numero: Número a ser buscado.
            
        Returns:
            True se encontrado, False caso contrário.
        """
        if not self.texto_completo:
            self.extrair_texto()
        
        # Remove caracteres não numéricos para comparação
        numero_limpo = re.sub(r'[^\d]', '', numero)
        texto_limpo = re.sub(r'[^\d]', '', self.texto_completo)
        
        return numero_limpo in texto_limpo


def verificar_boleto_por_conteudo(pdf_path: str, numero_boleto: str) -> bool:
    """
    Verifica se um PDF contém o número do boleto buscado.
    
    Args:
        pdf_path: Caminho para o arquivo PDF.
        numero_boleto: Número do boleto a ser verificado.
        
    Returns:
        True se o número for encontrado no conteúdo do PDF.
    """
    try:
        extractor = BoletoExtractor(pdf_path)
        return extractor.buscar_numero_no_texto(numero_boleto)
    except Exception as e:
        print(f"Erro ao verificar boleto: {e}")
        return False


def extrair_info_boleto(pdf_path: str) -> Dict[str, Optional[str]]:
    """
    Extrai informações de um boleto PDF.
    
    Args:
        pdf_path: Caminho para o arquivo PDF.
        
    Returns:
        Dicionário com as informações extraídas.
    """
    extractor = BoletoExtractor(pdf_path)
    return extractor.extrair_todos_dados()


# Exemplo de uso
if __name__ == "__main__":
    import sys
    
    if len(sys.argv) > 1:
        pdf_path = sys.argv[1]
        print(f"Extraindo dados do boleto: {pdf_path}")
        
        dados = extrair_info_boleto(pdf_path)
        
        print("\n=== Dados Extraídos ===")
        for chave, valor in dados.items():
            if valor:
                print(f"{chave}: {valor}")
    else:
        print("Uso: python pdf_utils.py <caminho_do_boleto.pdf>")
