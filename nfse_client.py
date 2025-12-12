"""
Módulo para consulta de NFSe (Nota Fiscal de Serviços Eletrônica) via API SEFIN.

Este módulo permite:
- Buscar a chave de acesso da NFSe através do ID DPS
- Consultar e obter o XML da NFSe
- Decodificar o XML compactado em GZip Base64
- Autenticação via certificado digital (.pfx/.p12)
"""

import requests
from requests_pkcs12 import Pkcs12Adapter
import base64
import gzip
import os
import sys
from configparser import ConfigParser
from typing import Tuple, Optional
from io import BytesIO


def get_config_path():
    """
    Retorna o caminho do arquivo de configuração.
    Prioridade:
    1. Variável de ambiente BUSCABOLETO_CONFIG
    2. config.ini no mesmo diretório do executável
    3. config.ini no diretório atual
    """
    # Verifica variável de ambiente
    env_path = os.environ.get('BUSCABOLETO_CONFIG')
    if env_path and os.path.exists(env_path):
        return env_path
    
    # Verifica no diretório do executável (para PyInstaller)
    if getattr(sys, 'frozen', False):
        exe_dir = os.path.dirname(sys.executable)
        exe_config = os.path.join(exe_dir, 'config.ini')
        if os.path.exists(exe_config):
            return exe_config
    
    # Verifica no diretório do script
    script_dir = os.path.dirname(os.path.abspath(__file__))
    script_config = os.path.join(script_dir, 'config.ini')
    if os.path.exists(script_config):
        return script_config
    
    # Retorna caminho padrão
    return 'config.ini'


class NFSeClient:
    """Cliente para consulta de NFSe via API SEFIN."""
    
    def __init__(self, config_path: str = None):
        """
        Inicializa o cliente NFSe com as configurações.
        
        Args:
            config_path: Caminho para o arquivo de configuração (opcional).
        """
        # Tenta carregar do arquivo de configuração
        if config_path is None:
            config_path = get_config_path()
        
        self.config = ConfigParser()
        config_loaded = False
        
        if os.path.exists(config_path):
            self.config.read(config_path, encoding='utf-8')
            config_loaded = True
        
        # Configurações de endpoints
        self.endpoint_iddps = self.config.get('ENDPOINTS', 'endpoint_nfse_iddps', fallback='').strip('"') if config_loaded else ''
        self.endpoint_chave_acesso = self.config.get('ENDPOINTS', 'endpoint_nfse_chave_acesso', fallback='').strip('"') if config_loaded else ''
        self.endpoint_pdf = self.config.get('ENDPOINTS', 'endpoint_nfse_pdf', fallback='').strip('"') if config_loaded else ''
        self.prefixo_iddps = self.config.get('ENDPOINTS', 'prefixo_iddps', fallback='').strip('"') if config_loaded else ''
        
        # Configurações de certificado digital
        self.certificado_path = self.config.get('CERTIFICADO', 'caminho', fallback='').strip('"') if config_loaded else ''
        self.certificado_senha = self.config.get('CERTIFICADO', 'senha', fallback='').strip('"') if config_loaded else ''
        
        # Configurações de busca
        self.timeout = self.config.getint('BUSCA', 'timeout', fallback=30) if config_loaded else 30
        
        # Pasta de downloads
        self.pasta_download = self.config.get('LOCAL', 'pasta_download', fallback='downloads').strip('"') if config_loaded else 'downloads'
        
        # Sessão HTTP com certificado (será criada sob demanda)
        self._session: Optional[requests.Session] = None
        
        # Criar pasta de download se não existir
        if not os.path.exists(self.pasta_download):
            os.makedirs(self.pasta_download)
    
    def _get_session(self) -> requests.Session:
        """
        Obtém ou cria uma sessão HTTP configurada com certificado digital.
        
        Returns:
            Sessão requests configurada.
        
        Raises:
            Exception: Se o certificado não estiver configurado ou não existir.
        """
        if self._session is not None:
            return self._session
        
        # Verifica se o certificado está configurado
        if not self.certificado_path:
            raise Exception(
                "Certificado digital não configurado.\n\n"
                "Configure a seção [CERTIFICADO] no config.ini:\n"
                "[CERTIFICADO]\n"
                "caminho = \"C:/caminho/para/certificado.pfx\"\n"
                "senha = \"sua_senha\""
            )
        
        # Verifica se o arquivo existe
        if not os.path.exists(self.certificado_path):
            raise Exception(
                f"Arquivo de certificado não encontrado:\n{self.certificado_path}\n\n"
                "Verifique o caminho configurado em [CERTIFICADO] no config.ini."
            )
        
        # Cria sessão com certificado PKCS12
        self._session = requests.Session()
        
        # Monta o adaptador com o certificado para os hosts dos endpoints
        # Extrai os hosts dos endpoints
        hosts_to_mount = set()
        
        if self.endpoint_iddps:
            # Extrai o base URL (https://host)
            from urllib.parse import urlparse
            parsed = urlparse(self.endpoint_iddps)
            base_url = f"{parsed.scheme}://{parsed.netloc}"
            hosts_to_mount.add(base_url)
        
        if self.endpoint_chave_acesso:
            from urllib.parse import urlparse
            parsed = urlparse(self.endpoint_chave_acesso)
            base_url = f"{parsed.scheme}://{parsed.netloc}"
            hosts_to_mount.add(base_url)
        
        if self.endpoint_pdf:
            from urllib.parse import urlparse
            parsed = urlparse(self.endpoint_pdf)
            base_url = f"{parsed.scheme}://{parsed.netloc}"
            hosts_to_mount.add(base_url)
        
        # Monta o adaptador PKCS12 para cada host
        for host in hosts_to_mount:
            self._session.mount(
                host,
                Pkcs12Adapter(
                    pkcs12_filename=self.certificado_path,
                    pkcs12_password=self.certificado_senha
                )
            )
        
        return self._session
    
    def verificar_certificado(self) -> Tuple[bool, str]:
        """
        Verifica se o certificado digital está configurado corretamente.
        
        Returns:
            Tupla com (sucesso: bool, mensagem: str)
        """
        if not self.certificado_path:
            return False, "Certificado não configurado no config.ini"
        
        if not os.path.exists(self.certificado_path):
            return False, f"Arquivo não encontrado: {self.certificado_path}"
        
        # Tenta carregar o certificado para verificar se a senha está correta
        try:
            from cryptography.hazmat.primitives.serialization import pkcs12
            from cryptography import x509
            
            with open(self.certificado_path, 'rb') as f:
                pfx_data = f.read()
            
            # Tenta carregar com a senha
            private_key, certificate, additional_certs = pkcs12.load_key_and_certificates(
                pfx_data,
                self.certificado_senha.encode() if self.certificado_senha else None
            )
            
            if certificate:
                # Extrai informações do certificado
                subject = certificate.subject.rfc4514_string()
                not_after = certificate.not_valid_after_utc
                
                return True, f"Certificado válido: {subject}\nVálido até: {not_after}"
            else:
                return False, "Certificado não contém dados válidos"
                
        except Exception as e:
            return False, f"Erro ao carregar certificado: {str(e)}"
    
    def formatar_numero_nfse(self, numero: str) -> str:
        """
        Formata o número da NFSe para 17 caracteres com zeros à esquerda.
        
        Args:
            numero: Número da NFSe (pode ser string ou número).
            
        Returns:
            Número formatado com 17 caracteres.
        
        Exemplo:
            "29" -> "00000000000000029"
        """
        # Remove caracteres não numéricos
        numero_limpo = ''.join(filter(str.isdigit, str(numero)))
        
        if not numero_limpo:
            return ""
        
        # Formata com zeros à esquerda para 17 caracteres
        return numero_limpo.zfill(17)
    
    def construir_id_dps(self, numero_nfse: str) -> str:
        """
        Constrói o ID DPS completo para consulta.
        
        Args:
            numero_nfse: Número da NFSe.
            
        Returns:
            ID DPS completo (prefixo + número formatado).
        
        Exemplo:
            "29" -> "DPS420540724779166800024900900000000000000029"
        """
        numero_formatado = self.formatar_numero_nfse(numero_nfse)
        if not numero_formatado:
            return ""
        
        return f"{self.prefixo_iddps}{numero_formatado}"
    
    def buscar_chave_acesso(self, numero_nfse: str) -> Tuple[bool, str, str]:
        """
        Busca a chave de acesso da NFSe através do ID DPS.
        
        Args:
            numero_nfse: Número da NFSe.
            
        Returns:
            Tupla com (sucesso: bool, chave_acesso ou mensagem_erro: str, id_dps: str)
        """
        id_dps = self.construir_id_dps(numero_nfse)
        
        if not id_dps:
            return False, "Número da NFSe inválido.", ""
        
        if not self.endpoint_iddps:
            return False, "Endpoint de busca DPS não configurado.", id_dps
        
        try:
            # Obtém sessão com certificado digital
            session = self._get_session()
            
            url = f"{self.endpoint_iddps}{id_dps}"
            
            response = session.get(url, timeout=self.timeout)
            
            if response.status_code == 200:
                dados = response.json()
                
                if 'chaveAcesso' in dados:
                    return True, dados['chaveAcesso'], id_dps
                else:
                    return False, f"Campo 'chaveAcesso' não encontrado na resposta. Resposta: {dados}", id_dps
            elif response.status_code == 404:
                return False, f"NFSe não encontrada para o ID: {id_dps}", id_dps
            elif response.status_code == 403:
                return False, f"Acesso negado (403). Verifique se o certificado digital está correto e válido.", id_dps
            elif response.status_code == 401:
                return False, f"Não autorizado (401). Certificado digital inválido ou expirado.", id_dps
            else:
                return False, f"Erro HTTP {response.status_code}: {response.text}", id_dps
                
        except requests.exceptions.Timeout:
            return False, f"Timeout ao consultar endpoint (>{self.timeout}s).", id_dps
        except requests.exceptions.ConnectionError as e:
            return False, f"Erro de conexão ao servidor SEFIN: {str(e)}", id_dps
        except requests.exceptions.JSONDecodeError:
            return False, f"Resposta inválida (não é JSON): {response.text[:200]}", id_dps
        except Exception as e:
            return False, f"Erro ao buscar chave de acesso: {str(e)}", id_dps
    
    def consultar_nfse(self, chave_acesso: str) -> Tuple[bool, str]:
        """
        Consulta a NFSe pela chave de acesso e obtém o XML.
        
        Args:
            chave_acesso: Chave de acesso da NFSe.
            
        Returns:
            Tupla com (sucesso: bool, xml_decodificado ou mensagem_erro: str)
        """
        if not chave_acesso:
            return False, "Chave de acesso não informada."
        
        if not self.endpoint_chave_acesso:
            return False, "Endpoint de consulta NFSe não configurado."
        
        try:
            # Obtém sessão com certificado digital
            session = self._get_session()
            
            url = f"{self.endpoint_chave_acesso}{chave_acesso}"
            
            response = session.get(url, timeout=self.timeout)
            
            if response.status_code == 200:
                dados = response.json()
                
                if 'nfseXmlGZipB64' in dados:
                    xml_gzip_b64 = dados['nfseXmlGZipB64']
                    xml_decodificado = self._decodificar_gzip_base64(xml_gzip_b64)
                    return True, xml_decodificado
                else:
                    return False, f"Campo 'nfseXmlGZipB64' não encontrado na resposta. Campos disponíveis: {list(dados.keys())}"
            elif response.status_code == 404:
                return False, f"NFSe não encontrada para a chave: {chave_acesso}"
            elif response.status_code == 403:
                return False, f"Acesso negado (403). Verifique se o certificado digital está correto e válido."
            elif response.status_code == 401:
                return False, f"Não autorizado (401). Certificado digital inválido ou expirado."
            else:
                return False, f"Erro HTTP {response.status_code}: {response.text}"
                
        except requests.exceptions.Timeout:
            return False, f"Timeout ao consultar NFSe (>{self.timeout}s)."
        except requests.exceptions.ConnectionError as e:
            return False, f"Erro de conexão ao servidor SEFIN: {str(e)}"
        except requests.exceptions.JSONDecodeError:
            return False, f"Resposta inválida (não é JSON): {response.text[:200]}"
        except Exception as e:
            return False, f"Erro ao consultar NFSe: {str(e)}"
    
    def _decodificar_gzip_base64(self, dados_codificados: str) -> str:
        """
        Decodifica dados em formato GZip Base64.
        
        Args:
            dados_codificados: String codificada em GZip + Base64.
            
        Returns:
            Dados decodificados como string.
        """
        try:
            # Decodifica Base64
            dados_gzip = base64.b64decode(dados_codificados)
            
            # Descompacta GZip
            with gzip.GzipFile(fileobj=BytesIO(dados_gzip)) as f:
                dados_decodificados = f.read()
            
            # Converte para string (UTF-8)
            return dados_decodificados.decode('utf-8')
        except Exception as e:
            raise Exception(f"Erro ao decodificar dados GZip Base64: {str(e)}")
    
    def buscar_xml_nfse(self, numero_nfse: str) -> Tuple[bool, str, dict]:
        """
        Busca o XML completo da NFSe pelo número.
        
        Este método combina as duas etapas:
        1. Buscar a chave de acesso via ID DPS
        2. Consultar o XML da NFSe via chave de acesso
        
        Args:
            numero_nfse: Número da NFSe.
            
        Returns:
            Tupla com (sucesso: bool, xml ou mensagem_erro: str, info: dict)
            O dict info contém: id_dps, chave_acesso (quando disponíveis)
        """
        info = {
            'numero_nfse': numero_nfse,
            'id_dps': '',
            'chave_acesso': ''
        }
        
        # Passo 1: Buscar chave de acesso
        sucesso, resultado, id_dps = self.buscar_chave_acesso(numero_nfse)
        info['id_dps'] = id_dps
        
        if not sucesso:
            return False, resultado, info
        
        chave_acesso = resultado
        info['chave_acesso'] = chave_acesso
        
        # Passo 2: Consultar NFSe
        sucesso, xml_ou_erro = self.consultar_nfse(chave_acesso)
        
        if not sucesso:
            return False, xml_ou_erro, info
        
        return True, xml_ou_erro, info
    
    def salvar_xml(self, xml_content: str, numero_nfse: str, pasta_destino: str = None) -> Tuple[bool, str]:
        """
        Salva o XML da NFSe em um arquivo.
        
        Args:
            xml_content: Conteúdo XML da NFSe.
            numero_nfse: Número da NFSe (usado no nome do arquivo).
            pasta_destino: Pasta onde salvar (usa pasta_download se não informada).
            
        Returns:
            Tupla com (sucesso: bool, caminho_arquivo ou mensagem_erro: str)
        """
        try:
            pasta = pasta_destino or self.pasta_download
            
            # Cria pasta se não existir
            if not os.path.exists(pasta):
                os.makedirs(pasta)
            
            # Nome do arquivo
            nome_arquivo = f"NFSe_{numero_nfse}.xml"
            caminho_completo = os.path.join(pasta, nome_arquivo)
            
            # Salva o arquivo
            with open(caminho_completo, 'w', encoding='utf-8') as f:
                f.write(xml_content)
            
            return True, caminho_completo
            
        except PermissionError:
            return False, "Erro de permissão ao salvar o arquivo."
        except Exception as e:
            return False, f"Erro ao salvar XML: {str(e)}"
    
    def buscar_e_salvar_xml_nfse(self, numero_nfse: str, pasta_destino: str = None) -> Tuple[bool, str, dict]:
        """
        Busca e salva o XML da NFSe em um único passo.
        
        Args:
            numero_nfse: Número da NFSe.
            pasta_destino: Pasta onde salvar (usa pasta_download se não informada).
            
        Returns:
            Tupla com (sucesso: bool, caminho_arquivo ou mensagem_erro: str, info: dict)
        """
        # Busca o XML
        sucesso, resultado, info = self.buscar_xml_nfse(numero_nfse)
        
        if not sucesso:
            return False, resultado, info
        
        xml_content = resultado
        
        # Salva o arquivo
        sucesso_salvar, caminho_ou_erro = self.salvar_xml(xml_content, numero_nfse, pasta_destino)
        
        if not sucesso_salvar:
            return False, caminho_ou_erro, info
        
        info['caminho_arquivo'] = caminho_ou_erro
        return True, caminho_ou_erro, info
    
    def baixar_pdf_nfse(self, chave_acesso: str, numero_nfse: str, pasta_destino: str = None) -> Tuple[bool, str]:
        """
        Baixa o PDF da NFSe diretamente da API pelo endpoint danfse.
        
        Args:
            chave_acesso: Chave de acesso da NFSe.
            numero_nfse: Número da NFSe (usado no nome do arquivo).
            pasta_destino: Pasta onde salvar (usa pasta_download se não informada).
            
        Returns:
            Tupla com (sucesso: bool, caminho_arquivo ou mensagem_erro: str)
        """
        if not chave_acesso:
            return False, "Chave de acesso não informada."
        
        if not self.endpoint_pdf:
            return False, "Endpoint de PDF (danfse) não configurado."
        
        try:
            # Obtém sessão com certificado digital
            session = self._get_session()
            
            url = f"{self.endpoint_pdf}{chave_acesso}"
            
            response = session.get(url, timeout=self.timeout)
            
            if response.status_code == 200:
                # Verifica se o conteúdo é PDF
                content_type = response.headers.get('Content-Type', '')
                if 'application/pdf' in content_type or response.content[:4] == b'%PDF':
                    # Salva o PDF
                    pasta = pasta_destino or self.pasta_download
                    
                    if not os.path.exists(pasta):
                        os.makedirs(pasta)
                    
                    nome_arquivo = f"NFSe_{numero_nfse}.pdf"
                    caminho_completo = os.path.join(pasta, nome_arquivo)
                    
                    with open(caminho_completo, 'wb') as f:
                        f.write(response.content)
                    
                    return True, caminho_completo
                else:
                    return False, f"Resposta não é PDF. Content-Type: {content_type}"
            elif response.status_code == 404:
                return False, f"PDF não encontrado para a chave: {chave_acesso}"
            elif response.status_code == 403:
                return False, f"Acesso negado (403). Verifique se o certificado digital está correto."
            elif response.status_code == 401:
                return False, f"Não autorizado (401). Certificado digital inválido ou expirado."
            else:
                return False, f"Erro HTTP {response.status_code}: {response.text[:200]}"
                
        except requests.exceptions.Timeout:
            return False, f"Timeout ao baixar PDF (>{self.timeout}s)."
        except requests.exceptions.ConnectionError as e:
            return False, f"Erro de conexão ao servidor: {str(e)}"
        except PermissionError:
            return False, "Erro de permissão ao salvar o arquivo PDF."
        except Exception as e:
            return False, f"Erro ao baixar PDF: {str(e)}"


def carregar_configuracao_endpoints(config_path: str = "config.ini") -> dict:
    """
    Carrega as configurações de endpoints do arquivo INI.
    
    Args:
        config_path: Caminho para o arquivo de configuração.
        
    Returns:
        Dicionário com as configurações de endpoints.
    """
    config = ConfigParser()
    config.read(config_path, encoding='utf-8')
    
    return {
        'endpoint_nfse_iddps': config.get('ENDPOINTS', 'endpoint_nfse_iddps', fallback='').strip('"'),
        'endpoint_nfse_chave_acesso': config.get('ENDPOINTS', 'endpoint_nfse_chave_acesso', fallback='').strip('"'),
        'endpoint_nfse_pdf': config.get('ENDPOINTS', 'endpoint_nfse_pdf', fallback='').strip('"'),
        'prefixo_iddps': config.get('ENDPOINTS', 'prefixo_iddps', fallback='').strip('"'),
    }
