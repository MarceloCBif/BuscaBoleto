# 🔍 Busca Boleto SFTP

Sistema para busca e download de boletos e notas fiscais de um servidor SFTP com interface gráfica.

## 📋 Funcionalidades

- ✅ Conexão automática com servidor SFTP
- ✅ Busca de boletos e NFs pelo número
- ✅ Busca por período de data
- ✅ Busca recursiva em subdiretórios (todas as filiais)
- ✅ Extração automática do nome do cliente do PDF
- ✅ Agrupamento de NF + Boleto por documento
- ✅ Seleção múltipla com checkboxes
- ✅ Download em arquivo ZIP
- ✅ Ordenação por colunas
- ✅ Interface gráfica amigável com Tkinter
- ✅ Restrição de acesso por IP (rede interna)
- ✅ **Consulta de XML da NFSe via API SEFIN** (Novo!)

## 📁 Estrutura do Projeto

```
BuscaBoleto/
├── main.py              # Arquivo principal para executar
├── interface.py         # Interface gráfica Tkinter
├── ftp_client.py        # Cliente SFTP para conexão e download
├── nfse_client.py       # Cliente para consulta de NFSe via API SEFIN
├── pdf_utils.py         # Utilitários para extração de dados do PDF
├── build_exe.py         # Script para gerar executável
├── config.ini           # Arquivo de configuração (não versionado)
├── config.example.ini   # Exemplo de configuração
├── requirements.txt     # Dependências do projeto
├── .gitignore           # Arquivos ignorados pelo Git
└── downloads/           # Pasta onde os boletos são salvos
```

## ⚙️ Configuração

### Opção 1: Arquivo de Configuração

1. Copie o arquivo de exemplo:
```bash
cp config.example.ini config.ini
```

2. Edite o arquivo `config.ini` com as informações do seu servidor SFTP:

```ini
[SFTP]
host = "sftp.seuservidor.com"
porta = 22
usuario = "seu_usuario"
senha = "sua_senha"
chave_privada = 
diretorio_remoto = "/caminho/para/boletos"
diretorio_remoto_nfs = "/caminho/para/nfs"

[LOCAL]
pasta_download = "downloads"

[BUSCA]
extensoes_permitidas = .pdf,.PDF
timeout = 30

[ENDPOINTS]
# Configurações para consulta de NFSe via API SEFIN
endpoint_nfse_iddps = "https://sefin.nfse.gov.br/SefinNacional/dps/"
endpoint_nfse_chave_acesso = "https://sefin.nfse.gov.br/SefinNacional/nfse/"
prefixo_iddps = "SEU_PREFIXO_IDDPS"
```

### Opção 2: Variáveis de Ambiente

Você pode configurar o sistema usando variáveis de ambiente (útil para CI/CD ou Docker):

| Variável | Descrição | Exemplo |
|----------|-----------|---------|
| `SFTP_HOST` | Endereço do servidor SFTP | `sftp.exemplo.com` |
| `SFTP_PORT` | Porta do servidor | `22` |
| `SFTP_USER` | Usuário para autenticação | `usuario` |
| `SFTP_PASSWORD` | Senha para autenticação | `senha123` |
| `SFTP_KEY_PATH` | Caminho para chave privada (opcional) | `/path/to/key` |
| `SFTP_BOLETO_DIR` | Diretório de boletos no servidor | `/boletos` |
| `SFTP_NF_DIR` | Diretório de NFs no servidor | `/nfs` |
| `DOWNLOAD_PATH` | Pasta local para downloads | `./downloads` |
| `BUSCABOLETO_CONFIG` | Caminho personalizado para config.ini | `/etc/app/config.ini` |

**Exemplo no PowerShell:**
```powershell
$env:SFTP_HOST = "sftp.exemplo.com"
$env:SFTP_USER = "usuario"
$env:SFTP_PASSWORD = "senha123"
python main.py
```

**Exemplo no Bash:**
```bash
export SFTP_HOST="sftp.exemplo.com"
export SFTP_USER="usuario"
export SFTP_PASSWORD="senha123"
python main.py
```

## 🚀 Instalação

### Pré-requisitos

- Python 3.8 ou superior
- Pip (gerenciador de pacotes Python)

### Passo a Passo

1. Clone o repositório:
```bash
git clone https://github.com/seu-usuario/BuscaBoleto.git
cd BuscaBoleto
```

2. Crie um ambiente virtual (recomendado):
```bash
python -m venv .venv

# Windows
.venv\Scripts\activate

# Linux/Mac
source .venv/bin/activate
```

3. Instale as dependências:
```bash
pip install -r requirements.txt
```

4. Configure o arquivo `config.ini` (veja seção Configuração)

5. Execute:
```bash
python main.py
```

## 🖥️ Como Usar

### Busca de Boletos e NFs (SFTP)

1. **Buscar**: Digite o número do documento no campo de busca e clique em "Buscar" (conexão automática)
2. **Filtrar por data**: Use os campos de data para buscar por período
3. **Selecionar**: Marque os checkboxes dos arquivos desejados
4. **Baixar**: Clique em "Baixar ZIP" para baixar os arquivos selecionados em um arquivo compactado

### Busca de XML NFSe (API SEFIN)

1. **Configurar**: Certifique-se de que a seção `[ENDPOINTS]` está configurada no `config.ini`
2. **Informar número**: Digite o número da NFSe no campo "Número da NFSe"
3. **Buscar**: Clique em "📄 Buscar XML NFSe"
4. **Resultado**: O sistema irá:
   - Consultar o ID DPS para obter a Chave de Acesso
   - Consultar a NFSe para obter o XML
   - Decodificar e salvar o arquivo XML na pasta de downloads
5. **Abrir**: Após o download, você pode abrir o arquivo XML diretamente

## 📦 Gerando Executável

Para gerar um executável `.exe` standalone:

```bash
python build_exe.py
```

O executável será gerado em `dist/BuscaBoleto.exe`.

**Importante:** Para distribuir o executável, inclua o arquivo `config.ini` na mesma pasta do `.exe`.

## 🔒 Segurança

- O arquivo `config.ini` contém credenciais sensíveis e **não deve ser versionado**
- Use variáveis de ambiente em ambientes de produção
- O sistema possui restrição de IP para funcionar apenas na rede interna (192.168.112.xxx)

## 📝 Dependências

```
paramiko>=3.0.0
pdfplumber>=0.10.0
requests>=2.28.0
pyinstaller>=6.0.0
```

## 🔧 Módulos

### ftp_client.py

Classe `SFTPClient` responsável por:
- Conectar/desconectar do servidor SFTP via SSH
- Autenticação por senha ou chave privada RSA
- Listar arquivos (simples e recursivo)
- Buscar boletos pelo número
- Baixar arquivos
- Auto-reconexão em caso de falha

### nfse_client.py

Classe `NFSeClient` responsável por:
- Construir o ID DPS a partir do número da NFSe
- Consultar a API SEFIN para obter a Chave de Acesso
- Consultar a API SEFIN para obter o XML da NFSe
- Decodificar o XML compactado em GZip Base64
- Salvar o XML em arquivo local

**Fluxo de consulta NFSe:**
1. Usuário informa o número da NFSe (ex: 29)
2. Sistema monta o ID DPS: `prefixo + numero(17 dígitos)` = `DPS420540724779166800024900900000000000000029`
3. Consulta endpoint `/dps/{id_dps}` para obter `chaveAcesso`
4. Consulta endpoint `/nfse/{chaveAcesso}` para obter `nfseXmlGZipB64`
5. Decodifica Base64, descompacta GZip
6. Salva o XML em arquivo

### pdf_utils.py

Classe `BoletoExtractor` responsável por:
- Extrair texto do PDF usando pdfplumber
- Identificar linha digitável com regex
- Extrair valor, vencimento, CNPJ
- Verificar se um número existe no boleto

### interface.py

Interface gráfica com:
- Campo de busca com máscara
- Filtros por data e filial
- **Campo para consulta de XML NFSe** (Novo!)
- Lista de resultados com checkboxes
- Agrupamento por documento (NF + Boleto)
- Botões de ação
- Barra de status e progresso
- Download em ZIP

## 🤝 Contribuindo

1. Faça um fork do projeto
2. Crie uma branch para sua feature (`git checkout -b feature/nova-feature`)
3. Commit suas mudanças (`git commit -am 'Adiciona nova feature'`)
4. Push para a branch (`git push origin feature/nova-feature`)
5. Abra um Pull Request

## ⚠️ Notas

- O sistema usa **SFTP** (SSH File Transfer Protocol), não FTP comum
- A porta padrão do SFTP é **22** (diferente do FTP que é 21)
- O sistema busca boletos pelo **nome do arquivo**, que deve conter o número do boleto
- Caracteres especiais são removidos durante a busca para melhor correspondência
- A busca é sempre recursiva em todos os subdiretórios
- Você pode usar chave privada RSA em vez de senha para autenticação

## 📄 Licença

Este projeto está sob a licença MIT. Veja o arquivo [LICENSE](LICENSE) para mais detalhes.
