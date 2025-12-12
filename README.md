# 🔍 Busca Boleto SFTP

Sistema para busca e download de boletos, notas fiscais e NFSe de um servidor SFTP com interface gráfica, incluindo integração com a API SEFIN para consulta de XML e PDF de NFSe.

## 📋 Funcionalidades

### Boletos e NFs (SFTP)
- ✅ Conexão automática com servidor SFTP
- ✅ Reconexão automática em caso de perda de conexão
- ✅ Busca de boletos e NFs pelo número
- ✅ Busca por período de data
- ✅ Busca recursiva em subdiretórios (todas as filiais)
- ✅ Extração automática do nome do cliente do PDF
- ✅ Agrupamento de NF + Boleto por documento
- ✅ Seleção múltipla com checkboxes
- ✅ Download em arquivo ZIP
- ✅ Ordenação por colunas
- ✅ Restrição de acesso por IP (rede interna)

### NFSe (API SEFIN)
- ✅ **Consulta automática de XML da NFSe** via API SEFIN
- ✅ **Download automático de PDF da NFSe** via API SEFIN
- ✅ **Autenticação com certificado digital** (.pfx)
- ✅ Extração do nome do cliente do XML (`<toma>/<xNome>`)
- ✅ Extração da data de emissão do XML (`<infDPS>/<dhEmi>`)
- ✅ Integração automática com resultados de boletos/NFs

### Interface
- ✅ Interface gráfica amigável com Tkinter
- ✅ Modal de reconexão automática
- ✅ Indicadores visuais de status de conexão
- ✅ Barra de progresso para operações longas

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
└── downloads/           # Pasta onde os arquivos são salvos
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
endpoint_nfse_pdf = "https://adn.nfse.gov.br/danfse/"
prefixo_iddps = "SEU_PREFIXO_IDDPS"

[CERTIFICADO]
# Certificado digital para autenticação na API SEFIN
caminho = "C:/caminho/para/certificado.pfx"
senha = "senha_do_certificado"
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

1. **Buscar por número**: Digite o número do documento no campo de busca e clique em "Buscar"
2. **Buscar por data**: Use os campos de data para buscar por período
3. **Selecionar**: Marque os checkboxes dos arquivos desejados
4. **Baixar**: Clique em "Baixar Selecionado(s)" para baixar os arquivos
5. **ZIP automático**: Múltiplos arquivos são baixados em um arquivo ZIP

### Busca Integrada de XML/PDF NFSe

Ao buscar boletos e NFs, o sistema **automaticamente**:

1. Identifica os números dos documentos encontrados
2. Consulta a API SEFIN para obter o XML da NFSe correspondente
3. Baixa o PDF oficial da NFSe via API
4. Adiciona os arquivos XML e PDF na lista de resultados
5. Permite download junto com boletos/NFs em um único ZIP

### Detalhes da Integração NFSe

O sistema utiliza certificado digital (.pfx) para autenticação na API SEFIN:

1. **Construção do ID DPS**: `prefixo + numero(17 dígitos)`
   - Exemplo: `DPS420540724779166800024900900000000000000029`
2. **Consulta Chave de Acesso**: Endpoint `/dps/{id_dps}`
3. **Consulta XML**: Endpoint `/nfse/{chaveAcesso}`
4. **Download PDF**: Endpoint `/danfse/{chaveAcesso}`
5. **Extração de dados**: Nome do cliente e data de emissão do XML

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
- **Certificado digital**: O arquivo `.pfx` deve ser protegido e não versionado

## 📝 Dependências

```
paramiko>=3.0.0
pdfplumber>=0.10.0
requests>=2.28.0
requests-pkcs12>=1.0.0
pyinstaller>=6.0.0
```

## 🔧 Módulos

### ftp_client.py

Classe `SFTPClient` responsável por:
- Conectar/desconectar do servidor SFTP via SSH
- Verificar status da conexão
- Autenticação por senha ou chave privada RSA
- Listar arquivos (simples e recursivo)
- Buscar boletos pelo número
- Baixar arquivos
- Extrair nome do cliente do PDF

### nfse_client.py

Classe `NFSeClient` responsável por:
- Autenticação com certificado digital PKCS12 (.pfx)
- Construir o ID DPS a partir do número da NFSe
- Consultar a API SEFIN para obter a Chave de Acesso
- Consultar a API SEFIN para obter o XML da NFSe
- **Baixar PDF oficial da NFSe** via endpoint `/danfse/`
- Decodificar o XML compactado em GZip Base64
- Salvar XML e PDF em arquivos locais

**Fluxo de consulta NFSe:**
```
Número NFSe (ex: 29)
    ↓
ID DPS: prefixo + numero(17 dígitos)
    ↓
GET /dps/{id_dps} → chaveAcesso
    ↓
GET /nfse/{chaveAcesso} → XML (GZip+Base64)
    ↓
GET /danfse/{chaveAcesso} → PDF
    ↓
Salva arquivos em downloads/
```

### pdf_utils.py

Classe `BoletoExtractor` responsável por:
- Extrair texto do PDF usando pdfplumber
- Identificar linha digitável com regex
- Extrair valor, vencimento, CNPJ
- Verificar se um número existe no boleto

### interface.py

Interface gráfica com:
- Campo de busca com máscara
- Filtros por data
- Lista de resultados com checkboxes
- Agrupamento por documento (NF + Boleto + XML + PDF)
- Download em ZIP
- **Reconexão automática com modal** ao perder conexão
- Extração de nome do cliente do XML NFSe
- Barra de status e progresso

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
- **NFSe**: Requer certificado digital válido (.pfx) para consulta na API SEFIN
- **Reconexão**: O sistema detecta perda de conexão e reconecta automaticamente

## 📄 Licença

Este projeto está sob a licença MIT. Veja o arquivo [LICENSE](LICENSE) para mais detalhes.
