"""
Módulo para geração de PDF a partir do XML de NFSe usando ReportLab.

Este módulo permite:
- Parsear XML de NFSe
- Gerar PDF formatado com os dados da nota
"""

import os
import xml.etree.ElementTree as ET
from datetime import datetime
from typing import Dict, Any, Optional

from reportlab.lib.pagesizes import A4
from reportlab.lib.units import mm, cm
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.enums import TA_CENTER, TA_LEFT, TA_RIGHT
from reportlab.lib.colors import HexColor, black, white
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle
from reportlab.platypus import HRFlowable


class NFSePDFGenerator:
    """Gerador de PDF para NFSe a partir de XML."""
    
    def __init__(self):
        """Inicializa o gerador de PDF."""
        self.styles = getSampleStyleSheet()
        self._criar_estilos_personalizados()
    
    def _criar_estilos_personalizados(self):
        """Cria estilos personalizados para o PDF."""
        # Título principal
        self.styles.add(ParagraphStyle(
            name='TituloPrincipal',
            parent=self.styles['Heading1'],
            fontSize=16,
            alignment=TA_CENTER,
            spaceAfter=10,
            textColor=HexColor('#1a237e')
        ))
        
        # Subtítulo
        self.styles.add(ParagraphStyle(
            name='Subtitulo',
            parent=self.styles['Heading2'],
            fontSize=12,
            alignment=TA_CENTER,
            spaceAfter=5,
            textColor=HexColor('#303f9f')
        ))
        
        # Título de seção
        self.styles.add(ParagraphStyle(
            name='TituloSecao',
            parent=self.styles['Heading3'],
            fontSize=10,
            alignment=TA_LEFT,
            spaceBefore=10,
            spaceAfter=5,
            textColor=HexColor('#1565c0'),
            backColor=HexColor('#e3f2fd'),
            borderPadding=5
        ))
        
        # Campo label
        self.styles.add(ParagraphStyle(
            name='CampoLabel',
            parent=self.styles['Normal'],
            fontSize=7,
            textColor=HexColor('#666666'),
            leading=9
        ))
        
        # Campo valor
        self.styles.add(ParagraphStyle(
            name='CampoValor',
            parent=self.styles['Normal'],
            fontSize=9,
            textColor=black,
            leading=11
        ))
        
        # Valor monetário
        self.styles.add(ParagraphStyle(
            name='ValorMonetario',
            parent=self.styles['Normal'],
            fontSize=10,
            alignment=TA_RIGHT,
            textColor=HexColor('#2e7d32'),
            fontName='Helvetica-Bold'
        ))
        
        # Rodapé
        self.styles.add(ParagraphStyle(
            name='Rodape',
            parent=self.styles['Normal'],
            fontSize=7,
            alignment=TA_CENTER,
            textColor=HexColor('#999999')
        ))
    
    def _get_tag_value(self, element: ET.Element, tag_name: str, default: str = '-') -> str:
        """
        Obtém o valor de uma tag do XML, buscando com ou sem namespace.
        
        Args:
            element: Elemento raiz ou pai para busca.
            tag_name: Nome da tag a buscar.
            default: Valor padrão se não encontrar.
            
        Returns:
            Valor da tag ou valor padrão.
        """
        # Tenta sem namespace
        elem = element.find(f'.//{tag_name}')
        if elem is not None and elem.text:
            return elem.text.strip()
        
        # Tenta com qualquer namespace
        for child in element.iter():
            local_name = child.tag.split('}')[-1] if '}' in child.tag else child.tag
            if local_name == tag_name and child.text:
                return child.text.strip()
        
        return default
    
    def _get_element(self, root: ET.Element, tag_name: str) -> Optional[ET.Element]:
        """
        Obtém um elemento do XML, buscando com ou sem namespace.
        
        Args:
            root: Elemento raiz para busca.
            tag_name: Nome da tag a buscar.
            
        Returns:
            Elemento encontrado ou None.
        """
        # Tenta sem namespace
        elem = root.find(f'.//{tag_name}')
        if elem is not None:
            return elem
        
        # Tenta com qualquer namespace
        for child in root.iter():
            local_name = child.tag.split('}')[-1] if '}' in child.tag else child.tag
            if local_name == tag_name:
                return child
        
        return None
    
    def _extrair_dados_xml(self, caminho_xml: str) -> Dict[str, Any]:
        """
        Extrai os dados relevantes do XML da NFSe.
        
        Args:
            caminho_xml: Caminho do arquivo XML.
            
        Returns:
            Dicionário com os dados extraídos.
        """
        dados = {
            # Dados da NFSe
            'numero_nfse': '-',
            'serie': '-',
            'codigo_verificacao': '-',
            'data_emissao': '-',
            'competencia': '-',
            'natureza_operacao': '-',
            'regime_tributacao': '-',
            
            # Prestador
            'prestador_cnpj': '-',
            'prestador_im': '-',
            'prestador_razao': '-',
            'prestador_fantasia': '-',
            'prestador_endereco': '-',
            'prestador_cidade': '-',
            'prestador_uf': '-',
            'prestador_cep': '-',
            'prestador_telefone': '-',
            'prestador_email': '-',
            
            # Tomador
            'tomador_cpf_cnpj': '-',
            'tomador_razao': '-',
            'tomador_endereco': '-',
            'tomador_cidade': '-',
            'tomador_uf': '-',
            'tomador_cep': '-',
            'tomador_telefone': '-',
            'tomador_email': '-',
            
            # Serviço
            'codigo_servico': '-',
            'discriminacao': '-',
            'cnae': '-',
            
            # Valores
            'valor_servicos': '0,00',
            'valor_deducoes': '0,00',
            'valor_pis': '0,00',
            'valor_cofins': '0,00',
            'valor_inss': '0,00',
            'valor_ir': '0,00',
            'valor_csll': '0,00',
            'valor_iss': '0,00',
            'aliquota_iss': '0,00',
            'valor_liquido': '0,00',
            'base_calculo': '0,00',
            'iss_retido': 'Não',
        }
        
        try:
            tree = ET.parse(caminho_xml)
            root = tree.getroot()
            
            # Dados da NFSe
            dados['numero_nfse'] = self._get_tag_value(root, 'nNFSe', dados['numero_nfse'])
            if dados['numero_nfse'] == '-':
                dados['numero_nfse'] = self._get_tag_value(root, 'Numero', dados['numero_nfse'])
            
            dados['serie'] = self._get_tag_value(root, 'Serie', dados['serie'])
            dados['codigo_verificacao'] = self._get_tag_value(root, 'CodigoVerificacao', dados['codigo_verificacao'])
            if dados['codigo_verificacao'] == '-':
                dados['codigo_verificacao'] = self._get_tag_value(root, 'cLocVerif', dados['codigo_verificacao'])
            
            # Data de emissão
            data_emissao = self._get_tag_value(root, 'dhEmi', '-')
            if data_emissao == '-':
                data_emissao = self._get_tag_value(root, 'DataEmissao', '-')
            if data_emissao != '-':
                try:
                    # Tenta diferentes formatos de data
                    if 'T' in data_emissao:
                        dt = datetime.fromisoformat(data_emissao.replace('Z', '+00:00').split('+')[0])
                        dados['data_emissao'] = dt.strftime('%d/%m/%Y %H:%M')
                    else:
                        dados['data_emissao'] = data_emissao
                except:
                    dados['data_emissao'] = data_emissao
            
            dados['competencia'] = self._get_tag_value(root, 'Competencia', dados['competencia'])
            if dados['competencia'] == '-':
                dados['competencia'] = self._get_tag_value(root, 'cCompet', dados['competencia'])
            
            # Prestador
            prestador = self._get_element(root, 'Prestador')
            if prestador is None:
                prestador = self._get_element(root, 'emit')
            if prestador is None:
                prestador = self._get_element(root, 'prest')
            
            if prestador is not None:
                dados['prestador_cnpj'] = self._get_tag_value(prestador, 'CNPJ', dados['prestador_cnpj'])
                if dados['prestador_cnpj'] == '-':
                    dados['prestador_cnpj'] = self._get_tag_value(prestador, 'Cnpj', dados['prestador_cnpj'])
                
                dados['prestador_im'] = self._get_tag_value(prestador, 'InscricaoMunicipal', dados['prestador_im'])
                if dados['prestador_im'] == '-':
                    dados['prestador_im'] = self._get_tag_value(prestador, 'IM', dados['prestador_im'])
                
                dados['prestador_razao'] = self._get_tag_value(prestador, 'RazaoSocial', dados['prestador_razao'])
                if dados['prestador_razao'] == '-':
                    dados['prestador_razao'] = self._get_tag_value(prestador, 'xNome', dados['prestador_razao'])
                
                dados['prestador_fantasia'] = self._get_tag_value(prestador, 'NomeFantasia', dados['prestador_fantasia'])
                if dados['prestador_fantasia'] == '-':
                    dados['prestador_fantasia'] = self._get_tag_value(prestador, 'xFant', dados['prestador_fantasia'])
                
                # Endereço do prestador
                end_prest = self._get_element(prestador, 'Endereco')
                if end_prest is None:
                    end_prest = self._get_element(prestador, 'enderPrest')
                if end_prest is None:
                    end_prest = self._get_element(prestador, 'enderEmit')
                
                if end_prest is not None:
                    logradouro = self._get_tag_value(end_prest, 'Logradouro', '')
                    if logradouro == '':
                        logradouro = self._get_tag_value(end_prest, 'xLgr', '')
                    numero = self._get_tag_value(end_prest, 'Numero', '')
                    if numero == '':
                        numero = self._get_tag_value(end_prest, 'nro', '')
                    bairro = self._get_tag_value(end_prest, 'Bairro', '')
                    if bairro == '':
                        bairro = self._get_tag_value(end_prest, 'xBairro', '')
                    
                    partes = [p for p in [logradouro, numero, bairro] if p]
                    dados['prestador_endereco'] = ', '.join(partes) if partes else '-'
                    
                    dados['prestador_cidade'] = self._get_tag_value(end_prest, 'CodigoMunicipio', '')
                    if dados['prestador_cidade'] == '':
                        dados['prestador_cidade'] = self._get_tag_value(end_prest, 'xMun', '-')
                    
                    dados['prestador_uf'] = self._get_tag_value(end_prest, 'Uf', dados['prestador_uf'])
                    if dados['prestador_uf'] == '-':
                        dados['prestador_uf'] = self._get_tag_value(end_prest, 'UF', dados['prestador_uf'])
                    
                    dados['prestador_cep'] = self._get_tag_value(end_prest, 'Cep', dados['prestador_cep'])
                    if dados['prestador_cep'] == '-':
                        dados['prestador_cep'] = self._get_tag_value(end_prest, 'CEP', dados['prestador_cep'])
                
                dados['prestador_telefone'] = self._get_tag_value(prestador, 'Telefone', dados['prestador_telefone'])
                if dados['prestador_telefone'] == '-':
                    dados['prestador_telefone'] = self._get_tag_value(prestador, 'fone', dados['prestador_telefone'])
                
                dados['prestador_email'] = self._get_tag_value(prestador, 'Email', dados['prestador_email'])
                if dados['prestador_email'] == '-':
                    dados['prestador_email'] = self._get_tag_value(prestador, 'email', dados['prestador_email'])
            
            # Tomador
            tomador = self._get_element(root, 'Tomador')
            if tomador is None:
                tomador = self._get_element(root, 'toma')
            if tomador is None:
                tomador = self._get_element(root, 'dest')
            
            if tomador is not None:
                cpf_cnpj = self._get_tag_value(tomador, 'CNPJ', '-')
                if cpf_cnpj == '-':
                    cpf_cnpj = self._get_tag_value(tomador, 'Cnpj', '-')
                if cpf_cnpj == '-':
                    cpf_cnpj = self._get_tag_value(tomador, 'CPF', '-')
                if cpf_cnpj == '-':
                    cpf_cnpj = self._get_tag_value(tomador, 'Cpf', '-')
                dados['tomador_cpf_cnpj'] = cpf_cnpj
                
                dados['tomador_razao'] = self._get_tag_value(tomador, 'RazaoSocial', dados['tomador_razao'])
                if dados['tomador_razao'] == '-':
                    dados['tomador_razao'] = self._get_tag_value(tomador, 'xNome', dados['tomador_razao'])
                
                # Endereço do tomador
                end_toma = self._get_element(tomador, 'Endereco')
                if end_toma is None:
                    end_toma = self._get_element(tomador, 'enderToma')
                if end_toma is None:
                    end_toma = self._get_element(tomador, 'enderDest')
                
                if end_toma is not None:
                    logradouro = self._get_tag_value(end_toma, 'Logradouro', '')
                    if logradouro == '':
                        logradouro = self._get_tag_value(end_toma, 'xLgr', '')
                    numero = self._get_tag_value(end_toma, 'Numero', '')
                    if numero == '':
                        numero = self._get_tag_value(end_toma, 'nro', '')
                    bairro = self._get_tag_value(end_toma, 'Bairro', '')
                    if bairro == '':
                        bairro = self._get_tag_value(end_toma, 'xBairro', '')
                    
                    partes = [p for p in [logradouro, numero, bairro] if p]
                    dados['tomador_endereco'] = ', '.join(partes) if partes else '-'
                    
                    dados['tomador_cidade'] = self._get_tag_value(end_toma, 'CodigoMunicipio', '')
                    if dados['tomador_cidade'] == '':
                        dados['tomador_cidade'] = self._get_tag_value(end_toma, 'xMun', '-')
                    
                    dados['tomador_uf'] = self._get_tag_value(end_toma, 'Uf', dados['tomador_uf'])
                    if dados['tomador_uf'] == '-':
                        dados['tomador_uf'] = self._get_tag_value(end_toma, 'UF', dados['tomador_uf'])
                    
                    dados['tomador_cep'] = self._get_tag_value(end_toma, 'Cep', dados['tomador_cep'])
                    if dados['tomador_cep'] == '-':
                        dados['tomador_cep'] = self._get_tag_value(end_toma, 'CEP', dados['tomador_cep'])
                
                dados['tomador_telefone'] = self._get_tag_value(tomador, 'Telefone', dados['tomador_telefone'])
                if dados['tomador_telefone'] == '-':
                    dados['tomador_telefone'] = self._get_tag_value(tomador, 'fone', dados['tomador_telefone'])
                
                dados['tomador_email'] = self._get_tag_value(tomador, 'Email', dados['tomador_email'])
                if dados['tomador_email'] == '-':
                    dados['tomador_email'] = self._get_tag_value(tomador, 'email', dados['tomador_email'])
            
            # Serviço
            servico = self._get_element(root, 'Servico')
            if servico is None:
                servico = self._get_element(root, 'serv')
            if servico is None:
                servico = self._get_element(root, 'DPS')
            if servico is None:
                servico = root  # Usa o root como fallback
            
            dados['codigo_servico'] = self._get_tag_value(servico, 'ItemListaServico', dados['codigo_servico'])
            if dados['codigo_servico'] == '-':
                dados['codigo_servico'] = self._get_tag_value(servico, 'cServ', dados['codigo_servico'])
            
            dados['discriminacao'] = self._get_tag_value(servico, 'Discriminacao', dados['discriminacao'])
            if dados['discriminacao'] == '-':
                dados['discriminacao'] = self._get_tag_value(servico, 'xDescServ', dados['discriminacao'])
            if dados['discriminacao'] == '-':
                dados['discriminacao'] = self._get_tag_value(root, 'xDescServ', dados['discriminacao'])
            
            dados['cnae'] = self._get_tag_value(servico, 'CodigoCnae', dados['cnae'])
            if dados['cnae'] == '-':
                dados['cnae'] = self._get_tag_value(servico, 'CNAE', dados['cnae'])
            
            # Valores
            valores = self._get_element(root, 'Valores')
            if valores is None:
                valores = self._get_element(root, 'vServPrest')
            if valores is None:
                valores = servico  # Usa serviço como fallback
            if valores is None:
                valores = root  # Usa o root como fallback
            
            def formatar_valor(valor_str: str) -> str:
                """Formata valor numérico para exibição."""
                try:
                    valor = float(valor_str.replace(',', '.'))
                    return f"{valor:,.2f}".replace(',', 'X').replace('.', ',').replace('X', '.')
                except:
                    return valor_str
            
            dados['valor_servicos'] = formatar_valor(self._get_tag_value(valores, 'ValorServicos', '0'))
            if dados['valor_servicos'] == '0':
                dados['valor_servicos'] = formatar_valor(self._get_tag_value(valores, 'vServPrest', '0'))
            if dados['valor_servicos'] == '0':
                dados['valor_servicos'] = formatar_valor(self._get_tag_value(root, 'vServPrest', '0'))
            if dados['valor_servicos'] == '0':
                dados['valor_servicos'] = formatar_valor(self._get_tag_value(root, 'vLiq', '0'))
            
            dados['valor_liquido'] = formatar_valor(self._get_tag_value(valores, 'ValorLiquidoNfse', '0'))
            if dados['valor_liquido'] == '0':
                dados['valor_liquido'] = formatar_valor(self._get_tag_value(root, 'vLiq', '0'))
            if dados['valor_liquido'] == '0':
                dados['valor_liquido'] = dados['valor_servicos']
            
            dados['valor_deducoes'] = formatar_valor(self._get_tag_value(valores, 'ValorDeducoes', '0'))
            dados['valor_pis'] = formatar_valor(self._get_tag_value(valores, 'ValorPis', '0'))
            dados['valor_cofins'] = formatar_valor(self._get_tag_value(valores, 'ValorCofins', '0'))
            dados['valor_inss'] = formatar_valor(self._get_tag_value(valores, 'ValorInss', '0'))
            dados['valor_ir'] = formatar_valor(self._get_tag_value(valores, 'ValorIr', '0'))
            dados['valor_csll'] = formatar_valor(self._get_tag_value(valores, 'ValorCsll', '0'))
            dados['valor_iss'] = formatar_valor(self._get_tag_value(valores, 'ValorIss', '0'))
            if dados['valor_iss'] == '0':
                dados['valor_iss'] = formatar_valor(self._get_tag_value(root, 'vISS', '0'))
            
            dados['base_calculo'] = formatar_valor(self._get_tag_value(valores, 'BaseCalculo', '0'))
            if dados['base_calculo'] == '0':
                dados['base_calculo'] = formatar_valor(self._get_tag_value(root, 'vBC', '0'))
            
            aliquota = self._get_tag_value(valores, 'Aliquota', '0')
            if aliquota == '0':
                aliquota = self._get_tag_value(root, 'pAliqAplic', '0')
            try:
                aliq_float = float(aliquota.replace(',', '.'))
                if aliq_float > 1:  # Se for percentual (ex: 5 ao invés de 0.05)
                    dados['aliquota_iss'] = f"{aliq_float:.2f}%"
                else:
                    dados['aliquota_iss'] = f"{aliq_float * 100:.2f}%"
            except:
                dados['aliquota_iss'] = aliquota
            
            iss_retido = self._get_tag_value(valores, 'IssRetido', '2')
            if iss_retido == '2':
                iss_retido = self._get_tag_value(root, 'indISSRet', '2')
            dados['iss_retido'] = 'Sim' if iss_retido == '1' else 'Não'
            
        except Exception as e:
            dados['erro'] = str(e)
        
        return dados
    
    def _criar_campo(self, label: str, valor: str) -> list:
        """
        Cria um par label/valor formatado.
        
        Args:
            label: Rótulo do campo.
            valor: Valor do campo.
            
        Returns:
            Lista com parágrafos formatados.
        """
        return [
            Paragraph(f"<b>{label}</b>", self.styles['CampoLabel']),
            Paragraph(valor if valor else '-', self.styles['CampoValor'])
        ]
    
    def gerar_pdf(self, caminho_xml: str, caminho_pdf: str = None) -> str:
        """
        Gera um PDF a partir do XML da NFSe.
        
        Args:
            caminho_xml: Caminho do arquivo XML.
            caminho_pdf: Caminho de saída do PDF (opcional, será gerado automaticamente).
            
        Returns:
            Caminho do arquivo PDF gerado.
        """
        # Define o caminho do PDF
        if caminho_pdf is None:
            base_name = os.path.splitext(caminho_xml)[0]
            caminho_pdf = f"{base_name}.pdf"
        
        # Extrai os dados do XML
        dados = self._extrair_dados_xml(caminho_xml)
        
        # Cria o documento PDF
        doc = SimpleDocTemplate(
            caminho_pdf,
            pagesize=A4,
            rightMargin=15*mm,
            leftMargin=15*mm,
            topMargin=15*mm,
            bottomMargin=15*mm
        )
        
        # Lista de elementos do documento
        elements = []
        
        # Título
        elements.append(Paragraph("NOTA FISCAL DE SERVIÇOS ELETRÔNICA - NFSe", self.styles['TituloPrincipal']))
        elements.append(Spacer(1, 5*mm))
        
        # Dados da NFSe
        nfse_data = [
            [Paragraph("<b>Número:</b>", self.styles['CampoLabel']),
             Paragraph(dados['numero_nfse'], self.styles['CampoValor']),
             Paragraph("<b>Série:</b>", self.styles['CampoLabel']),
             Paragraph(dados['serie'], self.styles['CampoValor']),
             Paragraph("<b>Data Emissão:</b>", self.styles['CampoLabel']),
             Paragraph(dados['data_emissao'], self.styles['CampoValor'])],
            [Paragraph("<b>Cód. Verificação:</b>", self.styles['CampoLabel']),
             Paragraph(dados['codigo_verificacao'], self.styles['CampoValor']),
             Paragraph("<b>Competência:</b>", self.styles['CampoLabel']),
             Paragraph(dados['competencia'], self.styles['CampoValor']),
             Paragraph("", self.styles['CampoLabel']),
             Paragraph("", self.styles['CampoValor'])],
        ]
        
        t = Table(nfse_data, colWidths=[55, 80, 40, 60, 55, 80])
        t.setStyle(TableStyle([
            ('VALIGN', (0, 0), (-1, -1), 'TOP'),
            ('TOPPADDING', (0, 0), (-1, -1), 2),
            ('BOTTOMPADDING', (0, 0), (-1, -1), 2),
            ('BOX', (0, 0), (-1, -1), 0.5, HexColor('#1565c0')),
            ('BACKGROUND', (0, 0), (-1, -1), HexColor('#f5f5f5')),
        ]))
        elements.append(t)
        elements.append(Spacer(1, 5*mm))
        
        # Seção: Prestador
        elements.append(Paragraph("PRESTADOR DE SERVIÇOS", self.styles['TituloSecao']))
        
        prest_data = [
            [Paragraph("<b>CNPJ:</b>", self.styles['CampoLabel']),
             Paragraph(dados['prestador_cnpj'], self.styles['CampoValor']),
             Paragraph("<b>Inscrição Municipal:</b>", self.styles['CampoLabel']),
             Paragraph(dados['prestador_im'], self.styles['CampoValor'])],
            [Paragraph("<b>Razão Social:</b>", self.styles['CampoLabel']),
             Paragraph(dados['prestador_razao'], self.styles['CampoValor']),
             Paragraph("<b>Nome Fantasia:</b>", self.styles['CampoLabel']),
             Paragraph(dados['prestador_fantasia'], self.styles['CampoValor'])],
            [Paragraph("<b>Endereço:</b>", self.styles['CampoLabel']),
             Paragraph(dados['prestador_endereco'], self.styles['CampoValor']),
             Paragraph("<b>Cidade/UF:</b>", self.styles['CampoLabel']),
             Paragraph(f"{dados['prestador_cidade']} / {dados['prestador_uf']}", self.styles['CampoValor'])],
            [Paragraph("<b>Telefone:</b>", self.styles['CampoLabel']),
             Paragraph(dados['prestador_telefone'], self.styles['CampoValor']),
             Paragraph("<b>Email:</b>", self.styles['CampoLabel']),
             Paragraph(dados['prestador_email'], self.styles['CampoValor'])],
        ]
        
        t = Table(prest_data, colWidths=[55, 130, 65, 130])
        t.setStyle(TableStyle([
            ('VALIGN', (0, 0), (-1, -1), 'TOP'),
            ('TOPPADDING', (0, 0), (-1, -1), 2),
            ('BOTTOMPADDING', (0, 0), (-1, -1), 2),
            ('BOX', (0, 0), (-1, -1), 0.5, HexColor('#cccccc')),
            ('LINEBELOW', (0, 0), (-1, -2), 0.25, HexColor('#eeeeee')),
        ]))
        elements.append(t)
        elements.append(Spacer(1, 3*mm))
        
        # Seção: Tomador
        elements.append(Paragraph("TOMADOR DE SERVIÇOS", self.styles['TituloSecao']))
        
        toma_data = [
            [Paragraph("<b>CPF/CNPJ:</b>", self.styles['CampoLabel']),
             Paragraph(dados['tomador_cpf_cnpj'], self.styles['CampoValor']),
             Paragraph("<b>Razão Social:</b>", self.styles['CampoLabel']),
             Paragraph(dados['tomador_razao'], self.styles['CampoValor'])],
            [Paragraph("<b>Endereço:</b>", self.styles['CampoLabel']),
             Paragraph(dados['tomador_endereco'], self.styles['CampoValor']),
             Paragraph("<b>Cidade/UF:</b>", self.styles['CampoLabel']),
             Paragraph(f"{dados['tomador_cidade']} / {dados['tomador_uf']}", self.styles['CampoValor'])],
            [Paragraph("<b>Telefone:</b>", self.styles['CampoLabel']),
             Paragraph(dados['tomador_telefone'], self.styles['CampoValor']),
             Paragraph("<b>Email:</b>", self.styles['CampoLabel']),
             Paragraph(dados['tomador_email'], self.styles['CampoValor'])],
        ]
        
        t = Table(toma_data, colWidths=[55, 130, 65, 130])
        t.setStyle(TableStyle([
            ('VALIGN', (0, 0), (-1, -1), 'TOP'),
            ('TOPPADDING', (0, 0), (-1, -1), 2),
            ('BOTTOMPADDING', (0, 0), (-1, -1), 2),
            ('BOX', (0, 0), (-1, -1), 0.5, HexColor('#cccccc')),
            ('LINEBELOW', (0, 0), (-1, -2), 0.25, HexColor('#eeeeee')),
        ]))
        elements.append(t)
        elements.append(Spacer(1, 3*mm))
        
        # Seção: Serviço
        elements.append(Paragraph("DISCRIMINAÇÃO DOS SERVIÇOS", self.styles['TituloSecao']))
        
        serv_data = [
            [Paragraph("<b>Código Serviço:</b>", self.styles['CampoLabel']),
             Paragraph(dados['codigo_servico'], self.styles['CampoValor']),
             Paragraph("<b>CNAE:</b>", self.styles['CampoLabel']),
             Paragraph(dados['cnae'], self.styles['CampoValor'])],
        ]
        
        t = Table(serv_data, colWidths=[55, 130, 55, 140])
        t.setStyle(TableStyle([
            ('VALIGN', (0, 0), (-1, -1), 'TOP'),
            ('TOPPADDING', (0, 0), (-1, -1), 2),
            ('BOTTOMPADDING', (0, 0), (-1, -1), 2),
        ]))
        elements.append(t)
        
        # Discriminação (texto longo)
        elements.append(Paragraph("<b>Discriminação:</b>", self.styles['CampoLabel']))
        discriminacao_text = dados['discriminacao'].replace('\n', '<br/>') if dados['discriminacao'] else '-'
        elements.append(Paragraph(discriminacao_text, self.styles['CampoValor']))
        elements.append(Spacer(1, 3*mm))
        
        # Seção: Valores
        elements.append(Paragraph("VALORES E TRIBUTOS", self.styles['TituloSecao']))
        
        valor_data = [
            [Paragraph("<b>Valor dos Serviços:</b>", self.styles['CampoLabel']),
             Paragraph(f"R$ {dados['valor_servicos']}", self.styles['ValorMonetario']),
             Paragraph("<b>Deduções:</b>", self.styles['CampoLabel']),
             Paragraph(f"R$ {dados['valor_deducoes']}", self.styles['CampoValor']),
             Paragraph("<b>Base de Cálculo:</b>", self.styles['CampoLabel']),
             Paragraph(f"R$ {dados['base_calculo']}", self.styles['CampoValor'])],
            [Paragraph("<b>PIS:</b>", self.styles['CampoLabel']),
             Paragraph(f"R$ {dados['valor_pis']}", self.styles['CampoValor']),
             Paragraph("<b>COFINS:</b>", self.styles['CampoLabel']),
             Paragraph(f"R$ {dados['valor_cofins']}", self.styles['CampoValor']),
             Paragraph("<b>INSS:</b>", self.styles['CampoLabel']),
             Paragraph(f"R$ {dados['valor_inss']}", self.styles['CampoValor'])],
            [Paragraph("<b>IR:</b>", self.styles['CampoLabel']),
             Paragraph(f"R$ {dados['valor_ir']}", self.styles['CampoValor']),
             Paragraph("<b>CSLL:</b>", self.styles['CampoLabel']),
             Paragraph(f"R$ {dados['valor_csll']}", self.styles['CampoValor']),
             Paragraph("<b>ISS Retido:</b>", self.styles['CampoLabel']),
             Paragraph(dados['iss_retido'], self.styles['CampoValor'])],
            [Paragraph("<b>Alíquota ISS:</b>", self.styles['CampoLabel']),
             Paragraph(dados['aliquota_iss'], self.styles['CampoValor']),
             Paragraph("<b>Valor ISS:</b>", self.styles['CampoLabel']),
             Paragraph(f"R$ {dados['valor_iss']}", self.styles['CampoValor']),
             Paragraph("", self.styles['CampoLabel']),
             Paragraph("", self.styles['CampoValor'])],
        ]
        
        t = Table(valor_data, colWidths=[55, 65, 50, 65, 55, 80])
        t.setStyle(TableStyle([
            ('VALIGN', (0, 0), (-1, -1), 'TOP'),
            ('TOPPADDING', (0, 0), (-1, -1), 2),
            ('BOTTOMPADDING', (0, 0), (-1, -1), 2),
            ('BOX', (0, 0), (-1, -1), 0.5, HexColor('#cccccc')),
            ('LINEBELOW', (0, 0), (-1, -2), 0.25, HexColor('#eeeeee')),
        ]))
        elements.append(t)
        elements.append(Spacer(1, 5*mm))
        
        # Valor Líquido em destaque
        valor_liquido_data = [
            [Paragraph("<b>VALOR LÍQUIDO DA NOTA:</b>", self.styles['CampoValor']),
             Paragraph(f"R$ {dados['valor_liquido']}", self.styles['ValorMonetario'])],
        ]
        
        t = Table(valor_liquido_data, colWidths=[280, 100])
        t.setStyle(TableStyle([
            ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
            ('TOPPADDING', (0, 0), (-1, -1), 5),
            ('BOTTOMPADDING', (0, 0), (-1, -1), 5),
            ('BOX', (0, 0), (-1, -1), 1, HexColor('#1565c0')),
            ('BACKGROUND', (0, 0), (-1, -1), HexColor('#e3f2fd')),
        ]))
        elements.append(t)
        elements.append(Spacer(1, 10*mm))
        
        # Rodapé
        elements.append(HRFlowable(width="100%", thickness=0.5, color=HexColor('#cccccc')))
        elements.append(Spacer(1, 2*mm))
        elements.append(Paragraph(
            f"Documento gerado em {datetime.now().strftime('%d/%m/%Y às %H:%M')} | "
            f"Fonte: XML da NFSe",
            self.styles['Rodape']
        ))
        
        # Gera o PDF
        doc.build(elements)
        
        return caminho_pdf


def gerar_pdf_nfse(caminho_xml: str, caminho_pdf: str = None) -> str:
    """
    Função utilitária para gerar PDF a partir de XML de NFSe.
    
    Args:
        caminho_xml: Caminho do arquivo XML.
        caminho_pdf: Caminho de saída do PDF (opcional).
        
    Returns:
        Caminho do arquivo PDF gerado.
    """
    generator = NFSePDFGenerator()
    return generator.gerar_pdf(caminho_xml, caminho_pdf)
