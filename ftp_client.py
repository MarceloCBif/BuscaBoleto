"""
Módulo de conexão SFTP para busca e download de boletos.
"""

import paramiko
import stat
import os
import sys
import re
from datetime import datetime
from configparser import ConfigParser
from typing import List, Optional, Tuple


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


class SFTPClient:
    """Cliente SFTP para busca e download de boletos."""
    
    def __init__(self, config_path: str = None):
        """
        Inicializa o cliente SFTP com as configurações.
        
        As configurações podem vir de:
        1. Variáveis de ambiente (prioridade)
        2. Arquivo de configuração (config.ini)
        
        Args:
            config_path: Caminho para o arquivo de configuração (opcional).
        
        Variáveis de ambiente suportadas:
            SFTP_HOST: Endereço do servidor SFTP
            SFTP_PORT: Porta do servidor (padrão: 22)
            SFTP_USER: Usuário para autenticação
            SFTP_PASSWORD: Senha para autenticação
            SFTP_KEY_PATH: Caminho para chave privada (opcional)
            SFTP_BOLETO_DIR: Diretório de boletos no servidor
            SFTP_NF_DIR: Diretório de NFs no servidor
            DOWNLOAD_PATH: Pasta local para downloads
        """
        # Tenta carregar do arquivo de configuração
        if config_path is None:
            config_path = get_config_path()
        
        self.config = ConfigParser()
        config_loaded = False
        
        if os.path.exists(config_path):
            self.config.read(config_path, encoding='utf-8')
            config_loaded = True
        
        # Configurações SFTP (variáveis de ambiente têm prioridade)
        self.host = os.environ.get('SFTP_HOST') or (
            self.config.get('SFTP', 'host', fallback='').strip('"') if config_loaded else ''
        )
        self.porta = int(os.environ.get('SFTP_PORT', 0)) or (
            self.config.getint('SFTP', 'porta', fallback=22) if config_loaded else 22
        )
        self.usuario = os.environ.get('SFTP_USER') or (
            self.config.get('SFTP', 'usuario', fallback='').strip('"') if config_loaded else ''
        )
        self.senha = os.environ.get('SFTP_PASSWORD') or (
            self.config.get('SFTP', 'senha', fallback='').strip('"') if config_loaded else ''
        )
        self.diretorio_remoto = os.environ.get('SFTP_BOLETO_DIR') or (
            self.config.get('SFTP', 'diretorio_remoto', fallback='').strip('"') if config_loaded else ''
        )
        self.diretorio_remoto_nfs = os.environ.get('SFTP_NF_DIR') or (
            self.config.get('SFTP', 'diretorio_remoto_nfs', fallback='').strip('"') if config_loaded else ''
        )
        
        # Chave privada (opcional)
        self.chave_privada = os.environ.get('SFTP_KEY_PATH') or (
            self.config.get('SFTP', 'chave_privada', fallback='').strip('"') if config_loaded else ''
        )
        
        # Configurações locais
        self.pasta_download = os.environ.get('DOWNLOAD_PATH') or (
            self.config.get('LOCAL', 'pasta_download', fallback='downloads').strip('"') if config_loaded else 'downloads'
        )
        
        # Configurações de busca
        extensoes = self.config.get('BUSCA', 'extensoes_permitidas', fallback='.pdf,.PDF') if config_loaded else '.pdf,.PDF'
        self.extensoes_permitidas = [ext.strip() for ext in extensoes.split(',')]
        self.timeout = self.config.getint('BUSCA', 'timeout', fallback=30) if config_loaded else 30
        
        self.ssh: Optional[paramiko.SSHClient] = None
        self.sftp: Optional[paramiko.SFTPClient] = None
        self.diretorio_atual = self.diretorio_remoto
        
        # Criar pasta de download se não existir
        if not os.path.exists(self.pasta_download):
            os.makedirs(self.pasta_download)
    
    def conectar(self) -> Tuple[bool, str]:
        """
        Estabelece conexão com o servidor SFTP.
        
        Returns:
            Tupla com (sucesso: bool, mensagem: str)
        """
        try:
            # Cria cliente SSH
            self.ssh = paramiko.SSHClient()
            self.ssh.set_missing_host_key_policy(paramiko.AutoAddPolicy())
            
            # Conecta usando senha ou chave privada
            if self.chave_privada and os.path.exists(self.chave_privada):
                chave = paramiko.RSAKey.from_private_key_file(self.chave_privada)
                self.ssh.connect(
                    hostname=self.host,
                    port=self.porta,
                    username=self.usuario,
                    pkey=chave,
                    timeout=self.timeout
                )
            else:
                self.ssh.connect(
                    hostname=self.host,
                    port=self.porta,
                    username=self.usuario,
                    password=self.senha,
                    timeout=self.timeout
                )
            
            # Abre sessão SFTP
            self.sftp = self.ssh.open_sftp()
            self.sftp.chdir(self.diretorio_remoto)
            self.diretorio_atual = self.diretorio_remoto
            
            return True, f"Conectado ao servidor SFTP {self.host}"
            
        except paramiko.AuthenticationException:
            return False, "Erro de autenticação: usuário ou senha inválidos"
        except paramiko.SSHException as e:
            return False, f"Erro SSH: {str(e)}"
        except Exception as e:
            return False, f"Erro ao conectar: {str(e)}"
    
    def desconectar(self):
        """Encerra a conexão com o servidor SFTP."""
        if self.sftp:
            try:
                self.sftp.close()
            except:
                pass
            self.sftp = None
        
        if self.ssh:
            try:
                self.ssh.close()
            except:
                pass
            self.ssh = None
    
    def verificar_conexao(self) -> bool:
        """
        Verifica se a conexão SFTP está ativa.
        
        Returns:
            True se a conexão está ativa, False caso contrário.
        """
        if not self.sftp or not self.ssh:
            return False
        
        try:
            # Tenta fazer uma operação simples para verificar a conexão
            self.sftp.getcwd()
            return True
        except:
            return False
    
    def reconectar(self) -> Tuple[bool, str]:
        """
        Reconecta ao servidor SFTP.
        
        Returns:
            Tupla com (sucesso: bool, mensagem: str)
        """
        self.desconectar()
        return self.conectar()
    
    def garantir_conexao(self) -> bool:
        """
        Garante que a conexão SFTP está ativa, reconectando se necessário.
        
        Returns:
            True se a conexão está ativa ou foi reconectada com sucesso.
        """
        if self.verificar_conexao():
            return True
        
        # Tenta reconectar
        sucesso, _ = self.reconectar()
        return sucesso
    
    def listar_arquivos(self, diretorio: str = None) -> List[str]:
        """
        Lista arquivos no diretório remoto.
        
        Args:
            diretorio: Diretório a ser listado (opcional).
            
        Returns:
            Lista de nomes de arquivos.
        """
        if not self.sftp:
            return []
        
        try:
            dir_listar = diretorio or self.sftp.getcwd() or self.diretorio_remoto
            
            arquivos = []
            for item in self.sftp.listdir_attr(dir_listar):
                # Verifica se é arquivo (não diretório)
                if not stat.S_ISDIR(item.st_mode):
                    # Filtrar apenas arquivos com extensões permitidas
                    if any(item.filename.endswith(ext) for ext in self.extensoes_permitidas):
                        arquivos.append(item.filename)
            
            return arquivos
        except Exception as e:
            print(f"Erro ao listar arquivos: {e}")
            return []
    
    def listar_arquivos_recursivo(self, diretorio: str = None) -> List[Tuple[str, str]]:
        """
        Lista arquivos recursivamente no diretório remoto.
        
        Args:
            diretorio: Diretório inicial (opcional).
            
        Returns:
            Lista de tuplas (caminho_completo, nome_arquivo).
        """
        if not self.sftp:
            return []
        
        arquivos_encontrados = []
        
        try:
            diretorio_atual = diretorio or self.sftp.getcwd() or self.diretorio_remoto
            
            for item in self.sftp.listdir_attr(diretorio_atual):
                caminho_completo = f"{diretorio_atual}/{item.filename}".replace("//", "/")
                
                if stat.S_ISDIR(item.st_mode):
                    # É um diretório, busca recursivamente
                    try:
                        arquivos_encontrados.extend(
                            self.listar_arquivos_recursivo(caminho_completo)
                        )
                    except PermissionError:
                        pass  # Sem permissão para acessar o diretório
                else:
                    # É um arquivo, verifica extensão
                    if any(item.filename.endswith(ext) for ext in self.extensoes_permitidas):
                        arquivos_encontrados.append((caminho_completo, item.filename))
                        
        except Exception as e:
            print(f"Erro ao listar recursivamente: {e}")
        
        return arquivos_encontrados
    
    def buscar_boleto(self, numero_boleto: str, busca_recursiva: bool = True, busca_literal: bool = False) -> List[Tuple[str, str, datetime]]:
        """
        Busca boletos pelo número no nome do arquivo.
        
        Args:
            numero_boleto: Número do boleto a ser buscado.
            busca_recursiva: Se True, busca em subdiretórios.
            busca_literal: Se True, busca o número exato no início do nome do arquivo.
                          Se False, busca se o número está contido em qualquer parte.
            
        Returns:
            Lista de tuplas (caminho_completo, nome_arquivo, data_modificacao) dos boletos encontrados.
        """
        if not self.sftp:
            return []
        
        # Limpa o número do boleto (remove caracteres especiais)
        numero_limpo = re.sub(r'[^\d]', '', numero_boleto)
        
        if not numero_limpo:
            return []
        
        # Obtém lista de arquivos com data
        if busca_recursiva:
            arquivos = self.listar_arquivos_com_data()
        else:
            dir_atual = self.sftp.getcwd() or self.diretorio_remoto
            arquivos = []
            for item in self.sftp.listdir_attr(dir_atual):
                if not stat.S_ISDIR(item.st_mode):
                    if any(item.filename.endswith(ext) for ext in self.extensoes_permitidas):
                        data_mod = datetime.fromtimestamp(item.st_mtime)
                        arquivos.append((f"{dir_atual}/{item.filename}", item.filename, data_mod))
        
        # Filtra arquivos que contêm o número do boleto no nome
        resultados = []
        for caminho, nome, data_mod in arquivos:
            # Remove extensão e caracteres especiais do nome para comparação
            nome_limpo = re.sub(r'[^\d]', '', nome)
            
            if busca_literal:
                # Busca literal: o número formatado (9 dígitos) deve aparecer após a filial (6 dígitos)
                # Exemplo: numero_limpo = "000005909" (9 dígitos)
                # nome do arquivo pode ser "010001000005909.pdf" (filial 010001 + numero 000005909)
                # Verifica se o nome tem pelo menos 15 dígitos (6 filial + 9 número)
                if len(nome_limpo) >= 15:
                    # Extrai os 9 últimos dígitos principais (após os 6 da filial)
                    numero_no_arquivo = nome_limpo[6:15]
                    if numero_no_arquivo == numero_limpo:
                        resultados.append((caminho, nome, data_mod))
                # Também aceita se o número aparecer em qualquer posição (para formatos diferentes)
                elif numero_limpo in nome_limpo:
                    resultados.append((caminho, nome, data_mod))
            else:
                # Busca parcial: verifica se o número está contido no nome do arquivo
                if numero_limpo in nome_limpo:
                    resultados.append((caminho, nome, data_mod))
        
        return resultados
    
    def listar_arquivos_com_data(self, diretorio: str = None) -> List[Tuple[str, str, datetime]]:
        """
        Lista arquivos recursivamente com suas datas de modificação.
        
        Args:
            diretorio: Diretório inicial (opcional).
            
        Returns:
            Lista de tuplas (caminho_completo, nome_arquivo, data_modificacao).
        """
        # Garante que a conexão está ativa
        if not self.garantir_conexao():
            return []
        
        arquivos_encontrados = []
        
        try:
            diretorio_atual = diretorio or self.sftp.getcwd() or self.diretorio_remoto
            
            for item in self.sftp.listdir_attr(diretorio_atual):
                caminho_completo = f"{diretorio_atual}/{item.filename}".replace("//", "/")
                
                if stat.S_ISDIR(item.st_mode):
                    # É um diretório, busca recursivamente
                    try:
                        arquivos_encontrados.extend(
                            self.listar_arquivos_com_data(caminho_completo)
                        )
                    except PermissionError:
                        pass  # Sem permissão para acessar o diretório
                    except Exception:
                        pass  # Ignora erros em subdiretórios
                else:
                    # É um arquivo, verifica extensão
                    if any(item.filename.endswith(ext) for ext in self.extensoes_permitidas):
                        # Converte timestamp para datetime
                        data_modificacao = datetime.fromtimestamp(item.st_mtime)
                        arquivos_encontrados.append((caminho_completo, item.filename, data_modificacao))
                        
        except Exception as e:
            # Se der erro de conexão, tenta reconectar e retorna lista vazia
            if "Garbage" in str(e) or "Socket" in str(e) or "EOF" in str(e):
                self.garantir_conexao()
            print(f"Erro ao listar recursivamente com data: {e}")
        
        return arquivos_encontrados
    
    def buscar_boleto_e_nf(self, numero_boleto: str, busca_recursiva: bool = True, busca_literal: bool = False) -> List[Tuple[str, str, datetime, str]]:
        """
        Busca boletos e notas fiscais pelo número no nome do arquivo.
        
        Args:
            numero_boleto: Número do boleto/NF a ser buscado.
            busca_recursiva: Se True, busca em subdiretórios.
            busca_literal: Se True, busca o número exato.
            
        Returns:
            Lista de tuplas (caminho_completo, nome_arquivo, data_modificacao, tipo) 
            onde tipo é 'BOLETO' ou 'NF'.
        """
        # Garante que a conexão está ativa
        if not self.garantir_conexao():
            return []
        
        # Limpa o número (remove caracteres especiais)
        numero_limpo = re.sub(r'[^\d]', '', numero_boleto)
        
        if not numero_limpo:
            return []
        
        resultados = []
        
        # Busca boletos
        boletos = self._buscar_em_diretorio(
            self.diretorio_remoto, numero_limpo, busca_recursiva, busca_literal, 'BOLETO'
        )
        resultados.extend(boletos)
        
        # Busca NFs se o diretório estiver configurado
        if self.diretorio_remoto_nfs:
            nfs = self._buscar_em_diretorio(
                self.diretorio_remoto_nfs, numero_limpo, busca_recursiva, busca_literal, 'NF'
            )
            resultados.extend(nfs)
        
        return resultados
    
    def _buscar_em_diretorio(self, diretorio: str, numero_limpo: str, busca_recursiva: bool, 
                              busca_literal: bool, tipo: str) -> List[Tuple[str, str, datetime, str]]:
        """
        Busca arquivos em um diretório específico.
        
        Args:
            diretorio: Diretório onde buscar.
            numero_limpo: Número já limpo para busca.
            busca_recursiva: Se True, busca em subdiretórios.
            busca_literal: Se True, busca o número exato.
            tipo: Tipo do arquivo ('BOLETO' ou 'NF').
            
        Returns:
            Lista de tuplas (caminho_completo, nome_arquivo, data_modificacao, tipo).
        """
        # Garante que a conexão está ativa
        if not self.garantir_conexao():
            return []
        
        try:
            # Obtém lista de arquivos com data
            if busca_recursiva:
                arquivos = self.listar_arquivos_com_data(diretorio)
            else:
                arquivos = []
                for item in self.sftp.listdir_attr(diretorio):
                    if not stat.S_ISDIR(item.st_mode):
                        if any(item.filename.endswith(ext) for ext in self.extensoes_permitidas):
                            data_mod = datetime.fromtimestamp(item.st_mtime)
                            arquivos.append((f"{diretorio}/{item.filename}", item.filename, data_mod))
            
            # Filtra arquivos que contêm o número no nome
            resultados = []
            for caminho, nome, data_mod in arquivos:
                # Remove extensão e caracteres especiais do nome para comparação
                nome_limpo = re.sub(r'[^\d]', '', nome)
                
                if busca_literal:
                    # Busca literal: o número formatado (9 dígitos) deve aparecer após a filial (6 dígitos)
                    if len(nome_limpo) >= 15:
                        numero_no_arquivo = nome_limpo[6:15]
                        if numero_no_arquivo == numero_limpo:
                            resultados.append((caminho, nome, data_mod, tipo))
                    elif numero_limpo in nome_limpo:
                        resultados.append((caminho, nome, data_mod, tipo))
                else:
                    # Busca parcial: verifica se o número está contido no nome do arquivo
                    if numero_limpo in nome_limpo:
                        resultados.append((caminho, nome, data_mod, tipo))
            
            return resultados
        except Exception as e:
            print(f"Erro ao buscar em {diretorio}: {e}")
            return []
    
    def buscar_por_data(self, data_inicio: datetime, data_fim: datetime) -> List[Tuple[str, str, datetime, str]]:
        """
        Busca arquivos por período de data de modificação em ambos diretórios.
        
        Args:
            data_inicio: Data inicial do período (inclusive).
            data_fim: Data final do período (inclusive).
            
        Returns:
            Lista de tuplas (caminho_completo, nome_arquivo, data_modificacao, tipo) dos arquivos encontrados.
        """
        # Garante que a conexão está ativa
        if not self.garantir_conexao():
            return []
        
        # Ajusta data_fim para incluir o dia inteiro (23:59:59)
        data_fim_ajustada = data_fim.replace(hour=23, minute=59, second=59)
        data_inicio_ajustada = data_inicio.replace(hour=0, minute=0, second=0)
        
        resultados = []
        
        # Busca boletos
        arquivos_boletos = self.listar_arquivos_com_data(self.diretorio_remoto)
        for caminho, nome, data_mod in arquivos_boletos:
            if data_inicio_ajustada <= data_mod <= data_fim_ajustada:
                resultados.append((caminho, nome, data_mod, 'BOLETO'))
        
        # Busca NFs se o diretório estiver configurado
        if self.diretorio_remoto_nfs:
            arquivos_nfs = self.listar_arquivos_com_data(self.diretorio_remoto_nfs)
            for caminho, nome, data_mod in arquivos_nfs:
                if data_inicio_ajustada <= data_mod <= data_fim_ajustada:
                    resultados.append((caminho, nome, data_mod, 'NF'))
        
        # Ordena por data (mais recentes primeiro)
        resultados.sort(key=lambda x: x[2], reverse=True)
        
        return resultados
    
    def baixar_boleto(self, caminho_remoto: str, nome_arquivo: str = None) -> Tuple[bool, str]:
        """
        Baixa um boleto do servidor SFTP.
        
        Args:
            caminho_remoto: Caminho completo do arquivo no servidor.
            nome_arquivo: Nome para salvar o arquivo (opcional).
            
        Returns:
            Tupla com (sucesso: bool, caminho_local ou mensagem_erro: str)
        """
        # Garante que a conexão está ativa
        if not self.garantir_conexao():
            return False, "Não conectado ao servidor SFTP"
        
        try:
            # Define nome do arquivo local
            if not nome_arquivo:
                nome_arquivo = os.path.basename(caminho_remoto)
            
            caminho_local = os.path.join(self.pasta_download, nome_arquivo)
            
            # Baixa o arquivo
            self.sftp.get(caminho_remoto, caminho_local)
            
            return True, caminho_local
            
        except PermissionError:
            return False, "Erro de permissão ao acessar o arquivo"
        except FileNotFoundError:
            return False, "Arquivo não encontrado no servidor"
        except Exception as e:
            return False, f"Erro ao baixar arquivo: {str(e)}"
    
    def extrair_cliente_do_pdf(self, caminho_remoto: str) -> str:
        """
        Extrai o nome do cliente de um arquivo PDF no servidor SFTP.
        
        Args:
            caminho_remoto: Caminho completo do arquivo no servidor.
            
        Returns:
            Nome do cliente ou string vazia se não encontrar.
        """
        import tempfile
        import pdfplumber
        
        # Garante que a conexão está ativa
        if not self.garantir_conexao():
            return ""
        
        try:
            # Cria arquivo temporário
            with tempfile.NamedTemporaryFile(suffix='.pdf', delete=False) as tmp:
                tmp_path = tmp.name
            
            # Baixa o arquivo para o temporário
            self.sftp.get(caminho_remoto, tmp_path)
            
            # Extrai texto do PDF
            cliente = ""
            with pdfplumber.open(tmp_path) as pdf:
                if pdf.pages:
                    texto = pdf.pages[0].extract_text() or ""
                    
                    # Tenta encontrar o nome do cliente
                    # Padrões comuns em boletos/NFs
                    linhas = texto.split('\n')
                    
                    for i, linha in enumerate(linhas):
                        linha_upper = linha.upper()
                        
                        # Padrão Boleto: "CLIENTE:" ou "SACADO:" seguido do nome
                        if 'CLIENTE:' in linha_upper or 'SACADO:' in linha_upper:
                            partes = linha.split(':', 1)
                            if len(partes) > 1 and partes[1].strip():
                                cliente = partes[1].strip()[:50]
                                break
                        
                        # Padrão: Linha após "CLIENTE" ou "SACADO" ou "PAGADOR"
                        if linha_upper.strip() in ['CLIENTE', 'SACADO', 'PAGADOR']:
                            if i + 1 < len(linhas) and linhas[i + 1].strip():
                                cliente = linhas[i + 1].strip()[:50]
                                break
                        
                        # Padrão NF: "NOME/RAZÃO SOCIAL" - nome está na próxima linha
                        # Mas a próxima linha NÃO pode ser um cabeçalho
                        if 'NOME/RAZÃO SOCIAL' in linha_upper or 'NOME/RAZAO SOCIAL' in linha_upper:
                            # Procura nas próximas linhas até encontrar um nome válido
                            for j in range(i + 1, min(i + 5, len(linhas))):
                                proxima_linha = linhas[j].strip()
                                proxima_upper = proxima_linha.upper()
                                
                                # Ignora linhas que são cabeçalhos ou vazias
                                if not proxima_linha:
                                    continue
                                if 'CNPJ' in proxima_upper and 'RAZÃO' in proxima_upper:
                                    continue
                                if 'CNPJ/CPF' in proxima_upper:
                                    continue
                                if proxima_upper.startswith('CNPJ') or proxima_upper.startswith('CPF'):
                                    continue
                                if 'DATA' in proxima_upper and 'EMISSÃO' in proxima_upper:
                                    continue
                                
                                # Encontrou uma linha válida - extrai o nome
                                import re
                                # Remove CNPJ/CPF e datas do final
                                nome_limpo = re.split(r'\d{2}[./]\d{2}[./]\d{2,4}|\d{2,3}[.]\d{3}[.]\d{3}', proxima_linha)[0]
                                nome_limpo = nome_limpo.strip()
                                
                                # Verifica se é um nome válido (não é só números ou muito curto)
                                if nome_limpo and len(nome_limpo) > 3 and not nome_limpo.replace(' ', '').isdigit():
                                    cliente = nome_limpo[:50]
                                    break
                            
                            if cliente:
                                break
                        
                        # Padrão: "RAZÃO SOCIAL:" ou "RAZAO SOCIAL:" (sem NOME/)
                        if ('RAZÃO SOCIAL:' in linha_upper or 'RAZAO SOCIAL:' in linha_upper) and 'NOME/' not in linha_upper:
                            partes = linha.split(':', 1)
                            if len(partes) > 1 and partes[1].strip():
                                cliente = partes[1].strip()[:50]
                                break
                        
                        # Padrão: "DESTINATÁRIO" ou "DESTINATARIO"
                        if 'DESTINATÁRIO' in linha_upper or 'DESTINATARIO' in linha_upper:
                            if ':' in linha:
                                partes = linha.split(':', 1)
                                if len(partes) > 1 and partes[1].strip():
                                    cliente = partes[1].strip()[:50]
                                    break
                            elif i + 1 < len(linhas) and linhas[i + 1].strip():
                                proxima = linhas[i + 1].strip()
                                # Ignora se for cabeçalho
                                if 'CNPJ' not in proxima.upper() and 'CPF' not in proxima.upper():
                                    cliente = proxima[:50]
                                    break
            
            # Remove arquivo temporário
            try:
                os.remove(tmp_path)
            except:
                pass
            
            return cliente
            
        except Exception as e:
            return ""


# Alias para manter compatibilidade
FTPClient = SFTPClient


def carregar_configuracao(config_path: str = "config.ini") -> dict:
    """
    Carrega as configurações do arquivo INI.
    
    Args:
        config_path: Caminho para o arquivo de configuração.
        
    Returns:
        Dicionário com as configurações.
    """
    config = ConfigParser()
    config.read(config_path, encoding='utf-8')
    
    return {
        'sftp': {
            'host': config.get('SFTP', 'host'),
            'porta': config.getint('SFTP', 'porta'),
            'usuario': config.get('SFTP', 'usuario'),
            'senha': config.get('SFTP', 'senha'),
            'chave_privada': config.get('SFTP', 'chave_privada', fallback=''),
            'diretorio_remoto': config.get('SFTP', 'diretorio_remoto'),
        },
        'local': {
            'pasta_download': config.get('LOCAL', 'pasta_download'),
        },
        'busca': {
            'extensoes_permitidas': config.get('BUSCA', 'extensoes_permitidas'),
            'timeout': config.getint('BUSCA', 'timeout'),
        }
    }
