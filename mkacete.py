# =================================================================================
# BLOCO 1: MÓDULOS E BIBLIOTECAS
# =================================================================================
# Este bloco importa todas as bibliotecas necessárias para o funcionamento do script.
# - pandas: Manipulação e análise de dados, especialmente para o arquivo Excel.
# - os: Interação com o sistema operacional (limpar tela, verificar arquivos).
# - time: Funções relacionadas a tempo (pausas e medição de performance).
# - sys: Acesso a parâmetros e funções do sistema (saída para o terminal).
# - unidecode: Normalização de texto, removendo acentuações e caracteres especiais.
# - difflib: Biblioteca para comparar sequências e fazer buscas 'fuzzy' (por similaridade).
# - json: Manipulação de dados no formato JSON, usado para o arquivo de configuração.
# - warnings: Para ignorar avisos de bibliotecas que não afetam o funcionamento.
# - colorama: Permite estilizar a saída do terminal com cores em diferentes sistemas operacionais.
# - pyperclip: Ferramenta para copiar texto para a área de transferência do sistema.

import pandas as pd
import os
import time
import sys
from unidecode import unidecode
from difflib import SequenceMatcher
import json
import warnings
import colorama
import pyperclip

# =================================================================================
# BLOCO 2: DADOS E CONSTANTES
# =================================================================================
# Este bloco contém os dicionários de dados que definem as frases de negativa
# e outras ferramentas de texto. Isso centraliza as frases em um único local,
# tornando a edição e a adição de novas opções mais fácil.

DADOS_RESTRICOES = {
    '01': {
        'nome': 'CARÊNCIA CONTRATUAL',
        'motivo': 'USUÁRIO EM PERÍODO DE CARÊNCIA',
        'fraseologia': """
Prezado(a) Sr(a). [_NM_BENEFICIARIO_]
NÚMERO DO PROTOCOLO: [_NU_PROTOCOLO_]

NO MOMENTO SUA SOLICITAÇÃO: {procedimento} FOI INDEFERIDA POR CARÊNCIA CONTRATUAL.

EXCETO CONSULTA E EXAMES DE URGÊNCIA SOLICITAMOS UMA NOVA ABERTURA DE DEMANDA PARA ESSE PROCEDIMENTO A PARTIR DA DATA: {data_disponivel}.

EM CASO DE DÚVIDAS, POR FAVOR, ENTRE EM CONTATO COM A CENTRAL DE ATENDIMENTO PELOS TELEFONES: 4090-1740, 0800 409 1740 OU 0800 463 4648.
"""
    },
    '02': {
        'nome': 'CONTRATO NÃO INICIADO EM PERÍODO DE VIGÊNCIA',
        'motivo': 'CONTRATO SÓ INICIA VALIDADE EM XX/XX/XXXX',
        'fraseologia': """
Prezado(a) Sr(a). [_NM_BENEFICIARIO_]
NÚMERO DO PROTOCOLO: [_NU_PROTOCOLO_]

A atuação da Operadora Hapvida está vinculada à regulação do Governo Federal, através da Agência Nacional de Saúde Suplementar - ANS, Autarquia Federal reguladora do referido setor de saúde. O contrato de plano de saúde individual ou familiar tem início de vigência a partir da data de assinatura da proposta de adesão, da assinatura do contrato ou da data de pagamento da mensalidade, o que ocorrer primeiro. No tocante aos contratos coletivos, a operadora de saúde e a pessoa jurídica contratante possuem liberdade para negociar o início da vigência, desde que até a data pactuada não haja qualquer pagamento. Verificando o contrato de Vossa Senhoria, identificou-se que o mesmo está com início de vigência programado para o dia {data_vigencia}. Dessa forma, apenas após tal data é que os serviços cobertos pela operadora poderão ser solicitados, observados os prazos de carências contratuais e legais.

EM CASO DE DÚVIDAS, POR FAVOR, ENTRE EM CONTATO COM A CENTRAL DE ATENDIMENTO PELOS TELEFONES: 4090-1740, 0800 409 1740 OU 0800 463 4648.
"""
    },
    '04': {
        'nome': 'EXCLUSÃO DE COBERTURA (AMBULATORIAL)',
        'motivo': 'PLANO CONTRATADO NAO COBERTO',
        'fraseologia': """
Prezado(a) Sr(a). [_NM_BENEFICIARIO_]
NÚMERO DO PROTOCOLO: [_NU_PROTOCOLO_]

A atuação da Operadora Hapvida está vinculada à regulação do Governo Federal, através da Agência Nacional de Saúde Suplementar - ANS, Autarquia Federal reguladora do referido setor de saúde. As coberturas são estabelecidas pela ANS, conforme previsto no Rol de Procedimentos e Eventos em Saúde, e determinadas pela Resolução Normativa n°. 465/2021, sob o amparo da Lei Federal n°. 9656/98. Após análise da solicitação de {procedimento01}, esta restou indeferida, pois verificou-se que o(a) beneficiário(a) é contratante de plano com segmentação exclusivamente AMBULATORIAL, registrado na Agência Nacional de Saúde Suplementar - ANS sob o nº XXX, sem direito à internação e/ou cirurgias. É válido salientar, que o plano de cobertura ambulatorial não abrange quaisquer atendimentos que necessitem de suporte em internação hospitalar, uma vez que a referida cobertura compreende, tão somente, consultas médicas em clínicas ou consultórios, exames, tratamento e demais procedimentos ambulatoriais, nos termos inciso I, do art. 12 da Lei Federal n° 9.656/1998 e do art. 18 da Resolução Normativa n°. 465/2021 da Agência Nacional de Saúde Suplementar. Dessa forma, o pedido para autorização do procedimento de {procedimento02}, não foi aprovado, por não se enquadrar em condições de cobertura contratualmente pactuadas. EM CASO DE DÚVIDAS, POR FAVOR, ENTRE EM CONTATO COM A CENTRAL DE ATENDIMENTO PELOS TELEFONES: 4090-1740, 0800 409 1740 OU 0800 463 4648.
"""
    },
    '08': {
        'nome': 'LIMITE CONTRATUAL',
        'motivo': 'PLANO CONVENCIONAL COM LIMITACAO CONTRATUAL',
        'fraseologia': """
Prezado(a) Sr(a). [_NM_BENEFICIARIO_]
NÚMERO DO PROTOCOLO: [_NU_PROTOCOLO_]

A atuação da Operadora Hapvida está vinculada à regulação do Governo Federal, através da Agência Nacional de Saúde Suplementar - ANS, Autarquia Federal reguladora do referido setor de saúde. As coberturas são estabelecidas pela ANS, conforme previsto no Rol de Procedimentos e Eventos em Saúde, e determinadas pela Resolução Normativa n°. 465/2021, sob o amparo da Lei Federal n°. 9656/98. O contrato de V.S.ª. trata-se de plano antigo, ou seja, foi comercializado antes da vigência da Lei 9.656/98. Tais planos possuem como característica a auto aplicabilidade de suas cláusulas, principalmente quanto aos eventos cobertos e excluídos. Sendo assim, a relação contratual estabelecida entre as partes é regida pelos termos exatos do contrato firmado. Em análise, pudemos verificar que o limite contratual foi excedido. Diante do exposto, a solicitação para autorização do procedimento {procedimento} acima mencionado não foi aprovada, por se tratar de pedido não coberto em contrato de plano de saúde não regulamentado, isto é, anterior à vigência da Lei nº 9.656/98.

EM CASO DE DÚVIDAS, POR FAVOR, ENTRE EM CONTATO COM A CENTRAL DE ATENDIMENTO PELOS TELEFONES: 4090-1740, 0800 409 1740 OU 0800 463 4648.
"""
    },
    '12': {
        'nome': 'INADIMPLÊNCIA CONTRATO COLETIVO',
        'motivo': 'PLANO EMPRESARIAL COM FATURA “EM ABERTO”',
        'fraseologia': """
PREZADO(A) SR(A). [_NM_BENEFICIARIO_]
NÚMERO DO PROTOCOLO: [_NU_PROTOCOLO_]

INFORMAMOS QUE, PARA DAR CONTINUIDADE AO ATENDIMENTO, É NECESSÁRIO QUE O(A) SENHOR(A) ENTRE EM CONTATO COM O SETOR DE RECURSOS HUMANOS (RH) DA SUA EMPRESA.

EM CASO DE DÚVIDAS, POR FAVOR, ENTRE EM CONTATO COM A CENTRAL DE ATENDIMENTO PELOS TELEFONES: 4090-1740, 0800 409 1740 OU 0800 463 4648.
"""
    },
    '16': {
        'nome': 'ÁREA DE ATUAÇÃO CONTRATUAL INCOMPATÍVEL',
        'motivo': 'REDE INCOMPATÍVEL COM O PLANO CONTRATADO',
        'fraseologia': """
PREZADO(A) SR(A). [_NM_BENEFICIARIO_]
NÚMERO DO PROTOCOLO: [_NU_PROTOCOLO_]

A atuação da Operadora Hapvida está vinculada à regulação do Governo Federal, através da Agência Nacional de Saúde Suplementar - ANS, Autarquia Federal reguladora do referido setor de saúde. As coberturas são estabelecidas pela ANS, conforme previsto no Rol de Procedimentos e Eventos em Saúde determinado pela Resolução Normativa n°. 465/2021, bem como, os critérios estabelecidos pela Lei Federal n°. 9.656/1998. Vossa Senhoria possui plano vinculado a esta operadora registrado na ANS sob o número XXX, com área de abrangência geográfica Grupo de XXX. Cumpre esclarecer que as operadoras de planos de assistência à saúde podem ofertar planos com área de abrangência Nacional, Estadual, de Grupo de Estados, Municipal ou de Grupo de Municípios, conforme esclarece item 4, do Anexo, da Resolução Normativa N° 543/2022, da Agência Nacional de Saúde Suplementar. Assim, vale enfatizar que o plano de saúde tem como área de atuação, tão somente, nos municípios ou estados albergados no referido tipo de plano contratado, com atendimento através dos médicos e prestadores indicados no Manual de Orientação do Beneficiário e Portal da Operadora, dentre os quais, não inclui a cidade de {cidade_estado}. Dessa forma, o pedido para autorização de {procedimento} acima mencionado, em atenção ao contrato celebrado, não foi aprovado, por não se enquadrar em condições de cobertura estabelecidas no instrumento contratual, haja vista estar fora da área de abrangência.

EM CASO DE DÚVIDAS, POR FAVOR, ENTRE EM CONTATO COM A CENTRAL DE ATENDIMENTO PELOS TELEFONES: 4090-1740, 0800 409 1740 OU 0800 463 4648.
"""
    },
    '24': {
        'nome': 'PERDA QUALIDADE DE BENEFICIÁRIO',
        'motivo': 'IMPOSSÍVEL AUTORIZAÇÃO SENHA, USUÁRIO NÃO ATIVO',
        'fraseologia': """
PREZADO(A) SR(A). [_NM_BENEFICIARIO_]
NÚMERO DO PROTOCOLO: [_NU_PROTOCOLO_]

A atuação da Operadora Hapvida está vinculada à regulação do Governo Federal, através da Agência Nacional de Saúde Suplementar - ANS, Autarquia Federal reguladora do referido setor de saúde. As coberturas são estabelecidas pela ANS, conforme previsto no Rol de Procedimentos e Eventos em Saúde, e determinadas pela Resolução Normativa n°. 465/2021, sob o amparo da Lei Federal n°. 9656/98. Após análise da solicitação de {procedimento}, esta restou indeferida, pois em conformidade com o Contrato, os serviços médicos prestados por esta operadora foram rescindidos, visto que vossa senhoria perdeu a qualidade de beneficiário.

SOLICITAMOS UMA NOVA ABERTURA COM CARTEIRINHA ATIVA. 

EM CASO DE DÚVIDAS, POR FAVOR, ENTRE EM CONTATO COM A CENTRAL DE ATENDIMENTO PELOS TELEFONES: 4090-1740, 0800 409 1740 OU 0800 463 4648.
"""
    },
    '26': {
        'nome': 'USUÁRIO É ATENDIDO POR OUTRA ASSISTÊNCIA',
        'motivo': 'USUÁRIO É ATENDIDO POR OUTRA ASSISTÊNCIA MÉDICA',
        'fraseologia': """
PREZADO(A) SR(A). [_NM_BENEFICIARIO_]
NÚMERO DO PROTOCOLO: [_NU_PROTOCOLO_]

NO MOMENTO SUA SOLICITAÇÃO FOI INDEFERIDA POR USUÁRIO É ATENDIDO POR OUTRA ASSISTÊNCIA.

A atuação da Operadora Hapvida está vinculada à regulação do Governo Federal, através da Agência Nacional de Saúde Suplementar - ANS, Autarquia Federal reguladora do referido setor de saúde. As coberturas são estabelecidas pela ANS, conforme previsto no Rol de Procedimentos e Eventos em Saúde, e determinadas pela Resolução Normativa n°. 465/2021, sob o amparo da Lei Federal n°. 9656/98. A Operadora deve garantir o acesso do beneficiário aos serviços e procedimentos definidos no Rol de Procedimentos e Eventos em Saúde da ANS, conforme os prazos da Resolução Normativa nº. 566/2022. A Resolução Normativa nº. 465/2021, que estabelece o Rol de Procedimentos e Eventos em Saúde, referência básica para cobertura assistencial mínima, em seu artigo 1º, §2º, assim dispõe: “A cobertura assistencial estabelecida por esta Resolução Normativa e seus anexos será obrigatória independente da circunstância e do local de ocorrência do evento que ensejar o atendimento, respeitadas as segmentações, a área de atuação e de abrangência, a rede de prestadores de serviços contratada, credenciada ou referenciada da operadora, os prazos de carência e a cobertura parcial temporária – CPT. Em relação à solicitação do procedimento de {procedimento}, verificamos que V.S.ª. está vinculado ao plano {nome_plano}, com prestação de serviço por outra Operadora de Saúde. Sendo assim, por não possuir previsão para prestadores dessa Operadora, referido procedimento fora indeferido.

EM CASO DE DÚVIDAS, POR FAVOR, ENTRE EM CONTATO COM A CENTRAL DE ATENDIMENTO PELOS TELEFONES: 4090-1740, 0800 409 1740 OU 0800 463 4648.
"""
    },
    '48': {
        'nome': 'PLANO AMBULATORIAL (SEM INTERNAÇÃO)',
        'motivo': 'PLANO DO USUARIO SOMENTE AMBULATORIAL',
        'fraseologia': """
PREZADO(A) SR(A). [_NM_BENEFICIARIO_]
NÚMERO DO PROTOCOLO: [_NU_PROTOCOLO_]

A atuação da Operadora Hapvida está vinculada à regulação do Governo Federal, através da Agência Nacional de Saúde Suplementar - ANS, Autarquia Federal reguladora do referido setor de saúde. As coberturas são estabelecidas pela ANS, conforme previsto no Rol de Procedimentos e Eventos em Saúde, e determinadas pela Resolução Normativa n°. 465/2021, sob o amparo da Lei Federal n°. 9656/98.Após análise da solicitação de {procedimento01}, esta restou indeferida, pois verificou-se que o(a) beneficiário(a) é contratante de plano com segmentação exclusivamente AMBULATORIAL, registrado na Agência Nacional de Saúde Suplementar - ANS sob o nº XXX, sem direito à internação e/ou cirurgias. É válido salientar, que o plano de cobertura ambulatorial não abrange quaisquer atendimentos que necessitem de suporte em internação hospitalar, uma vez que a referida cobertura compreende, tão somente, consultas médicas em clínicas ou consultórios, exames, tratamento e demais procedimentos ambulatoriais, nos termos inciso I, do art. 12 da Lei Federal n° 9.656/1998 e do art. 18 da Resolução Normativa n°. 465/2021 da Agência Nacional de Saúde Suplementar. Dessa forma, o pedido para autorização do procedimento de {procedimento02}, não foi aprovado, por não se enquadrar em condições de cobertura contratualmente pactuadas.

EM CASO DE DÚVIDAS, POR FAVOR, ENTRE EM CONTATO COM A CENTRAL DE ATENDIMENTO PELOS TELEFONES: 4090-1740, 0800 409 1740 OU 0800 463 4648.
"""
    },
    '49': {
        'nome': 'PLANO HOSPITALAR (SEM AMBULATORIAL)',
        'motivo': 'PLANO DO BENEFICIARIO SOMENTE HOSPITALAR',
        'fraseologia': """
PREZADO(A) SR(A). [_NM_BENEFICIARIO_]
NÚMERO DO PROTOCOLO: [_NU_PROTOCOLO_]

A atuação da Operadora Hapvida está vinculada à regulação do Governo Federal, através da Agência Nacional de Saúde Suplementar - ANS, Autarquia Federal reguladora do referido setor de saúde. As coberturas são estabelecidas pela ANS, conforme previsto no Rol de Procedimentos e Eventos em Saúde, e determinadas pela Resolução Normativa n°. 465/2021, sob o amparo da Lei Federal n°. 9656/98. Após análise da solicitação de {procedimento01}, esta restou indeferida, pois verificou-se que o beneficiário(a) é contratante de plano com segmentação exclusivamente HOSPITALAR, registrado na Agência Nacional de Saúde Suplementar - ANS sob o nº XXX, não incluindo atendimentos ambulatoriais para fins de diagnóstico, terapia ou recuperação. É válido salientar que, o plano de cobertura hospitalar compreende os atendimentos realizados em todas as modalidades de internação hospitalar e os atendimentos caracterizados como de urgência e emergência, nos termos inciso II, do art. 12 da Lei Federal n° 9.656/1998 e do art. 19 da Resolução Normativa n°. 465/2021 da Agência Nacional de Saúde Suplementar. Dessa forma, o pedido para autorização do procedimento de {procedimento02}, não foi aprovado, por não se enquadrar em condições de cobertura contratualmente pactuadas.

EM CASO DE DÚVIDAS, POR FAVOR, ENTRE EM CONTATO COM A CENTRAL DE ATENDIMENTO PELOS TELEFONES: 4090-1740, 0800 409 1740 OU 0800 463 4648.
"""
    }
}

FERRAMENTAS_TEXTO = {
    'REEMBOLSO': {
        'nome': 'ORIENTAÇÃO PARA REEMBOLSO',
        'fraseologia': """
PREZADO(A) [_NM_BENEFICIARIO_],

A SOLICITAÇÃO REGISTRADA NO PROTOCOLO [_NU_PROTOCOLO_] NÃO PODE SER TRATADA POR ESTE CANAL.

PARA PROSSEGUIR COM O PEDIDO DE REEMBOLSO, ENVIE A DOCUMENTAÇÃO DIRETAMENTE PARA O E-MAIL: REEMBOLSONDISPRJ@HAPVIDA.COM.BR

EM CASO DE DÚVIDAS, ENTRE EM CONTATO COM A NOSSA CENTRAL DE ATENDIMENTO PELOS NÚMEROS: 0800 463 4648, 4090 1740 OU 0800 409 1740.
"""
    }
}


# =================================================================================
# BLOCO 3: CLASSES PRINCIPAIS
# =================================================================================
# Este bloco define as classes que contêm a lógica central do sistema.
# - A classe 'Colors' gerencia o esquema de cores para o terminal.
# - A classe 'MecanismoBuscaAvancado' lida com o carregamento e a busca nos dados.

class Colors:
    """Define cores e estilos para estilizar o texto no terminal com o tema 'Batman'."""
    # Cores primárias do tema
    BATMAN_YELLOW = '\033[38;5;220m'  # Amarelo vibrante do logo do Batman
    GOTHAM_TEXT = '\033[38;5;250m'  # Cinza claro para o texto
    DARK_BLUE = '\033[38;5;20m'  # Azul escuro para bordas
    # Cores para feedback do sistema
    SUCCESS_GREEN = '\033[38;5;82m'
    ERROR_RED = '\033[38;5;196m'
    WARNING_YELLOW = '\033[38;5;220m'
    INFO_CYAN = '\033[38;5;45m'
    # Estilos
    BOLD = '\033[1m'
    UNDERLINE = '\033[4m'
    ENDC = '\033[0m'


class MecanismoBuscaAvancado:
    """
    Gerencia o carregamento de dados a partir de um arquivo Excel e a busca
    inteligente dentro das abas da planilha. Inclui funcionalidades de cache e
    busca por relevância e similaridade (fuzzy).
    """

    def __init__(self, nome_arquivo_excel):
        self.nome_arquivo_excel = nome_arquivo_excel
        self.dados_abas = {}
        self.cache_busca = {}
        self.estatisticas = {
            'total_buscas': 0,
            'tempo_medio_busca': 0,
            'resultados_encontrados': 0
        }
        self.config = self._carregar_configuracao()
        self._carregar_dados()

    def _carregar_configuracao(self):
        """Carrega as configurações do arquivo 'config.json' ou usa as padrão."""
        config_padrao = {
            'max_resultados': 100,
            'limiar_similaridade': 0.6,
            'habilitar_cache': True,
            'tamanho_cache': 1000,
            'habilitar_busca_fuzzy': True,
            'colunas_prioritarias': ['PRESTADOR', 'PROCEDIMENTOS', 'TUSS', 'AMB'],
            'pesos_colunas': {
                'PRESTADOR': 2.0,
                'PROCEDIMENTOS': 1.8,
                'TUSS': 1.5,
                'AMB': 1.5,
                'CNPJ': 1.2,
                'RAZAO SOCIAL': 1.3
            }
        }
        try:
            if os.path.exists('config.json'):
                with open('config.json', 'r', encoding='utf-8') as f:
                    config_personalizada = json.load(f)
                    config_padrao.update(config_personalizada)
        except Exception as e:
            print(f"{Colors.WARNING_YELLOW}Aviso: Erro ao carregar configuração personalizada: {e}{Colors.ENDC}")
        return config_padrao

    def _carregar_dados(self):
        """Carrega todos os dados do arquivo Excel para a memória."""
        if not os.path.exists(self.nome_arquivo_excel):
            print(
                f"{Colors.ERROR_RED}ERRO: Arquivo '{self.nome_arquivo_excel}' não detectado. Verifique a matriz de dados.{Colors.ENDC}")
            return
        try:
            inicio = time.time()
            self.dados_abas = pd.read_excel(
                self.nome_arquivo_excel,
                sheet_name=None,
                engine='openpyxl',
                na_filter=False
            )
            self._preparar_dados_busca()
            tempo_carregamento = time.time() - inicio
            print(f"{Colors.SUCCESS_GREEN}✓ Banco de dados online em {tempo_carregamento:.2f}s!{Colors.ENDC}")
            print(f"{Colors.INFO_CYAN}Total de abas carregadas: {len(self.dados_abas)}{Colors.ENDC}")
        except Exception as e:
            print(f"{Colors.ERROR_RED}Falha crítica ao carregar o arquivo: {e}{Colors.ENDC}")

    def _preparar_dados_busca(self):
        """Prepara os dados para busca, limpando e normalizando texto."""
        for nome_aba, df in self.dados_abas.items():
            df_limpo = df.dropna(how='all')
            df_limpo = df_limpo.fillna('')
            for col in df_limpo.columns:
                if df_limpo[col].dtype == 'object':
                    df_limpo[col] = df_limpo[col].astype(str).str.lower().apply(
                        lambda x: unidecode(x) if x else ''
                    )
            df_limpo['_TEXTO_BUSCA'] = df_limpo.apply(
                lambda row: ' '.join([str(val) for val in row.values if str(val).strip()]),
                axis=1
            )
            self.dados_abas[nome_aba] = df_limpo

    def buscar_avancada(self, nome_aba, termo, filtros=None, max_resultados=None):
        """Executa a busca por um termo em uma aba específica, usando algoritmos combinados."""
        if not self.dados_abas or nome_aba not in self.dados_abas:
            return pd.DataFrame()
        inicio = time.time()
        df = self.dados_abas[nome_aba].copy()
        max_resultados = max_resultados or self.config['max_resultados']
        chave_cache = f"{nome_aba}_{termo}_{str(filtros)}_{max_resultados}"
        if self.config['habilitar_cache'] and chave_cache in self.cache_busca:
            self.estatisticas['total_buscas'] += 1
            return self.cache_busca[chave_cache]
        if filtros:
            df = self._aplicar_filtros(df, filtros)
        resultados = self._busca_multi_algoritmo(df, termo, max_resultados)
        tempo_busca = time.time() - inicio
        self.estatisticas['total_buscas'] += 1
        self.estatisticas['tempo_medio_busca'] = (
                (self.estatisticas['tempo_medio_busca'] * (self.estatisticas['total_buscas'] - 1) + tempo_busca)
                / self.estatisticas['total_buscas']
        )
        self.estatisticas['resultados_encontrados'] += len(resultados)
        if self.config['habilitar_cache']:
            if len(self.cache_busca) >= self.config['tamanho_cache']:
                self.cache_busca.pop(next(iter(self.cache_busca)))
            self.cache_busca[chave_cache] = resultados
        return resultados

    def _busca_multi_algoritmo(self, df, termo, max_resultados):
        """Combina resultados de busca exata, por relevância e fuzzy para um resultado mais completo."""
        termo_limpo = unidecode(termo).lower().strip()
        resultados_exatos = self._busca_exata(df, termo_limpo)
        resultados_relevancia = self._busca_por_relevancia(df, termo_limpo, max_resultados)
        resultados_fuzzy = self._busca_fuzzy(df, termo_limpo, max_resultados)
        df_combinado = pd.concat([resultados_exatos, resultados_relevancia, resultados_fuzzy], axis=0)
        df_resultado = df_combinado.drop_duplicates(keep='first').head(max_resultados)
        return df_resultado

    def _busca_exata(self, df, termo):
        """Busca por correspondência exata do termo nas colunas."""
        mascara = df['_TEXTO_BUSCA'].str.contains(termo, na=False)
        return df[mascara]

    def _busca_por_relevancia(self, df, termo, max_resultados):
        """Busca e pontua os resultados com base na relevância e posição do termo."""
        resultados = []
        for _, row in df.iterrows():
            score = 0
            for col, peso in self.config['pesos_colunas'].items():
                if col in df.columns:
                    if termo in str(row[col]):
                        score += peso
                        if str(row[col]).startswith(termo):
                            score += 0.5
                        if f' {termo} ' in f' {str(row[col])} ':
                            score += 0.3
            if score > 0:
                resultados.append((score, row))
        resultados.sort(key=lambda x: x[0], reverse=True)
        return pd.DataFrame([row for _, row in resultados[:max_resultados]])

    def _busca_fuzzy(self, df, termo, max_resultados):
        """Busca por similaridade fonética ou de escrita (fuzzy)."""
        resultados = []
        for _, row in df.iterrows():
            melhor_similaridade = 0
            for col in self.config['colunas_prioritarias']:
                if col in df.columns and str(row[col]).strip():
                    similaridade = SequenceMatcher(None, termo, str(row[col])).ratio()
                    melhor_similaridade = max(melhor_similaridade, similaridade)
            if melhor_similaridade >= self.config['limiar_similaridade']:
                resultados.append((melhor_similaridade, row))
        resultados.sort(key=lambda x: x[0], reverse=True)
        return pd.DataFrame([row for _, row in resultados[:max_resultados]])

    def _aplicar_filtros(self, df, filtros):
        """Aplica filtros adicionais aos resultados da busca."""
        df_filtrado = df.copy()
        for coluna, valor in filtros.items():
            if coluna in df_filtrado.columns and valor:
                mascara = df_filtrado[coluna].str.contains(str(valor), na=False)
                df_filtrado = df_filtrado[mascara]
        return df_filtrado

    def exibir_resultados_avancados(self, resultados, termo, nome_aba, tempo_busca=None):
        """Exibe os resultados da busca de forma estilizada no terminal."""
        if resultados.empty:
            print(
                f"\n{Colors.WARNING_YELLOW}Nenhum resultado encontrado para '{termo}' no setor '{nome_aba}'.{Colors.ENDC}")
            return
        print(f"\n{Colors.INFO_CYAN}>>> Resultado da Busca Terminal <<<{Colors.ENDC}")
        print(f"{Colors.GOTHAM_TEXT}┌─ Termo de busca: {Colors.BOLD}{Colors.BATMAN_YELLOW}{termo}{Colors.ENDC}")
        print(f"{Colors.GOTHAM_TEXT}├─ Setor de dados: {Colors.BOLD}{Colors.BATMAN_YELLOW}{nome_aba}{Colors.ENDC}")
        print(
            f"{Colors.GOTHAM_TEXT}├─ Total de resultados: {Colors.BOLD}{Colors.BATMAN_YELLOW}{len(resultados)}{Colors.ENDC}")
        if tempo_busca:
            print(
                f"{Colors.GOTHAM_TEXT}└─ Tempo de execução: {Colors.BOLD}{Colors.BATMAN_YELLOW}{tempo_busca:.3f}s{Colors.ENDC}")
        print(f"{Colors.INFO_CYAN}-----------------------------------------{Colors.ENDC}")
        if 'infiltracao' in nome_aba.lower() or 'medicacao' in nome_aba.lower():
            self._exibir_resultados_por_tema(resultados)
        else:
            self._exibir_generico_futurista(resultados)
        print(f"\n{Colors.INFO_CYAN}-----------------------------------------{Colors.ENDC}")

    def _exibir_resultados_por_tema(self, resultados):
        """Agrupa e exibe resultados por categorias específicas."""
        temas = {
            "INFILTRAÇÃO / BLOQUEIO": ['infiltracao', 'infiltração', 'bloqueio'],
            "RETIRADA DE MEDICAMENTO": ['retirada', 'remocao', 'remover'],
        }
        df_restantes = resultados.copy()
        for titulo_tema, palavras_chave in temas.items():
            mask = df_restantes['_TEXTO_BUSCA'].str.contains('|'.join(palavras_chave), na=False)
            df_tema = df_restantes[mask]
            if not df_tema.empty:
                self._exibir_bloco_categoria(titulo_tema, df_tema)
            df_restantes = df_restantes[~mask]
        if not df_restantes.empty:
            self._exibir_bloco_categoria("Outros Resultados (Sem Categoria)", df_restantes)

    def _exibir_bloco_categoria(self, titulo, df):
        """Função auxiliar para exibir blocos de resultados temáticos."""
        if not df.empty:
            print(f"\n{Colors.BOLD}{Colors.INFO_CYAN}----- {titulo} ({len(df)}) -----{Colors.ENDC}")
            self._exibir_generico_futurista(df)

    def _exibir_generico_futurista(self, df):
        """Formata e exibe os resultados da busca de forma genérica e estilizada."""
        for i, (_, row) in enumerate(df.iterrows(), 1):
            print(
                f"\n{Colors.BOLD}{Colors.INFO_CYAN}═══════════════ [{i:03d}] Resultado ═══════════════{Colors.ENDC}")
            key_color = Colors.GOTHAM_TEXT
            value_color = Colors.BOLD + Colors.BATMAN_YELLOW
            colunas_prioritarias = ['PRESTADOR', 'ZONA/REGIÃO', 'CNPJ', 'CD PESSOA']
            for col in colunas_prioritarias:
                if col in row and str(row[col]).strip():
                    print(f"  {key_color}{str(col).upper()}:{Colors.ENDC} {value_color}{row[col]}{Colors.ENDC}")
            for col, val in row.items():
                if str(val).strip() and col != '_TEXTO_BUSCA' and col not in colunas_prioritarias:
                    print(f"  {key_color}{str(col).upper()}:{Colors.ENDC} {value_color}{val}{Colors.ENDC}")
            print(
                f"{Colors.BOLD}{Colors.INFO_CYAN}═══════════════════════════════════════════════════════{Colors.ENDC}")

    def mostrar_estatisticas(self):
        """Exibe as estatísticas de uso e performance do sistema."""
        print(f"\n{Colors.INFO_CYAN}>>> Relatório de Status do Sistema <<<{Colors.ENDC}")
        print(
            f"{Colors.GOTHAM_TEXT}Total de consultas: {Colors.BOLD}{self.estatisticas['total_buscas']}{Colors.ENDC}")
        print(
            f"{Colors.GOTHAM_TEXT}Tempo médio de resposta: {Colors.BOLD}{self.estatisticas['tempo_medio_busca']:.3f}s{Colors.ENDC}")
        print(
            f"{Colors.GOTHAM_TEXT}Total de resultados gerados: {Colors.BOLD}{self.estatisticas['resultados_encontrados']}{Colors.ENDC}")
        print(f"{Colors.GOTHAM_TEXT}Cache de memória: {Colors.BOLD}{len(self.cache_busca)}{Colors.ENDC}")
        print(
            f"{Colors.GOTHAM_TEXT}Status do cache: {Colors.BOLD}{'Ativo' if self.config['habilitar_cache'] else 'Inativo'}{Colors.ENDC}")
        print(f"{Colors.INFO_CYAN}-----------------------------------------{Colors.ENDC}")

    def limpar_cache(self):
        """Limpa o cache de busca para liberar memória."""
        self.cache_busca.clear()
        print(f"{Colors.SUCCESS_GREEN}✓ Cache de busca limpo com sucesso! {Colors.ENDC}")

    def salvar_configuracao(self):
        """Salva as configurações atuais do sistema em um arquivo JSON."""
        try:
            with open('config.json', 'w', encoding='utf-8') as f:
                json.dump(self.config, f, indent=2, ensure_ascii=False)
            print(f"{Colors.SUCCESS_GREEN}✓ Configuração do sistema salva em 'config.json'!{Colors.ENDC}")
        except Exception as e:
            print(f"{Colors.ERROR_RED}Falha ao salvar a configuração: {e}{Colors.ENDC}")


# =================================================================================
# BLOCO 4: FUNÇÕES DE UTILIDADE
# =================================================================================
# Funções simples que auxiliam na formatação do terminal e outras tarefas menores.

def center_text(text, width):
    """Função utilitária para centralizar texto no terminal."""
    return text.center(width)


# =================================================================================
# BLOCO 5: FUNÇÕES DE GERAÇÃO DE FRASES
# =================================================================================
# Funções dedicadas à criação de frases prontas para autorizações, negativas e
# reembolsos, com a opção de copiar o texto para a área de transferência.

def gerar_texto_reembolso():
    """Gera e copia a fraseologia para orientação de reembolso."""
    try:
        frase_reembolso = FERRAMENTAS_TEXTO['REEMBOLSO']['fraseologia']
        pyperclip.copy(frase_reembolso)
        print(f"\n{Colors.SUCCESS_GREEN}✓ Texto de Reembolso copiado para a área de transferência!{Colors.ENDC}")
        print(f"\n{Colors.BOLD}{Colors.INFO_CYAN}----------------------------------------{Colors.ENDC}")
        print(f"{Colors.GOTHAM_TEXT}{frase_reembolso}{Colors.ENDC}")
        print(f"{Colors.BOLD}{Colors.INFO_CYAN}----------------------------------------{Colors.ENDC}")
    except Exception as e:
        print(f"{Colors.ERROR_RED}Erro ao gerar texto de reembolso: {e}{Colors.ENDC}")


def gerar_fraseologia_positiva():
    """Gera e copia uma fraseologia de autorização para exames ou procedimentos."""
    try:
        try:
            num_procedimentos_str = input(
                f"{Colors.GOTHAM_TEXT}Quantos procedimentos serão autorizados? {Colors.ENDC}").strip()
            num_procedimentos = int(num_procedimentos_str)
            if num_procedimentos <= 0:
                print(f"{Colors.WARNING_YELLOW}O número de procedimentos deve ser maior que zero. Saindo.{Colors.ENDC}")
                return
        except ValueError:
            print(f"{Colors.ERROR_RED}Entrada inválida. Digite um número inteiro.{Colors.ENDC}")
            return

        fraseologia_base = """
PREZADO(A) SR(A). [_NM_BENEFICIARIO_],
NÚMERO DO PROTOCOLO: [_NU_PROTOCOLO_]

SUA SOLICITAÇÃO DE AUTORIZAÇÃO PARA EXAME FOI RECEBIDA COM OS SEGUINTES DADOS:
"""
        corpo_fraseologia = ""
        for i in range(num_procedimentos):
            procedimento = input(
                f"{Colors.GOTHAM_TEXT}Procedimento ({i + 1}/{num_procedimentos}): {Colors.ENDC}").strip().upper()
            senha = input(f"{Colors.GOTHAM_TEXT}Senha ({i + 1}/{num_procedimentos}): {Colors.ENDC}").strip()
            prestador = input(
                f"{Colors.GOTHAM_TEXT}Prestador ({i + 1}/{num_procedimentos}): {Colors.ENDC}").strip().upper()
            if not procedimento or not senha or not prestador:
                print(f"{Colors.WARNING_YELLOW}Todos os campos são obrigatórios. Saindo.{Colors.ENDC}")
                return
            corpo_fraseologia += f"\nPROCEDIMENTO: {procedimento}"
            corpo_fraseologia += f"\nSENHA: {senha}"
            corpo_fraseologia += f"\nPRESTADOR: {prestador}"
            if i < num_procedimentos - 1:
                corpo_fraseologia += "\n================="

        mensagem_contato = """
EM CASO DE DÚVIDAS, POR FAVOR, ENTRE EM CONTATO COM A CENTRAL DE ATENDIMENTO PELOS TELEFONES: 4090-1740, 0800 409 1740 OU 0800 463 4648.
"""
        frase_final = fraseologia_base + corpo_fraseologia + mensagem_contato
        frase_limpa = os.linesep.join([s.strip() for s in frase_final.splitlines() if s.strip()])
        pyperclip.copy(frase_limpa)
        print(
            f"\n{Colors.SUCCESS_GREEN}✓ Fraseologia de Autorização copiada para a área de transferência!{Colors.ENDC}")
        print(f"\n{Colors.BOLD}{Colors.INFO_CYAN}----------------------------------------{Colors.ENDC}")
        print(f"{Colors.GOTHAM_TEXT}{frase_limpa}{Colors.ENDC}")
        print(f"{Colors.BOLD}{Colors.INFO_CYAN}----------------------------------------{Colors.ENDC}")
    except Exception as e:
        print(f"{Colors.ERROR_RED}Erro ao gerar fraseologia positiva: {e}{Colors.ENDC}")


def gerar_fraseologia_negativa():
    """Gera e copia uma fraseologia de negativa com base em um código."""
    print(f"\n{Colors.BOLD}{Colors.INFO_CYAN}--- Ferramenta de Negativa ---{Colors.ENDC}")
    print(f"{Colors.GOTHAM_TEXT}Códigos de negativa disponíveis:{Colors.ENDC}")
    for codigo, info in DADOS_RESTRICOES.items():
        print(f"{Colors.WARNING_YELLOW}[{codigo}]{Colors.ENDC} - {info['nome']}")
    escolha_negativa = input(f"{Colors.GOTHAM_TEXT}Escolha o código de negativa: {Colors.ENDC}").strip()
    if escolha_negativa not in DADOS_RESTRICOES:
        print(f"{Colors.ERROR_RED}Código de negativa inválido.{Colors.ENDC}")
        return
    info_negativa = DADOS_RESTRICOES[escolha_negativa]
    frase_modelo = info_negativa['fraseologia']
    campos_a_preencher = {}
    campos_disponiveis = [
        ('procedimento', '{procedimento}'),
        ('procedimento01', '{procedimento01}'),
        ('procedimento02', '{procedimento02}'),
        ('data_disponivel', '{data_disponivel}'),
        ('data_vigencia', '{data_vigencia}'),
        ('cidade_estado', '{cidade_estado}'),
        ('nome_plano', '{nome_plano}')
    ]
    for nome_campo, placeholder in campos_disponiveis:
        if placeholder in frase_modelo:
            if nome_campo == 'data_disponivel' or nome_campo == 'data_vigencia':
                input_prompt = f"{Colors.GOTHAM_TEXT}{nome_campo.replace('_', ' ').capitalize()} (dd/mm/aaaa): {Colors.ENDC}"
            elif nome_campo == 'cidade_estado':
                input_prompt = f"{Colors.GOTHAM_TEXT}Cidade ou Estado: {Colors.ENDC}"
            else:
                input_prompt = f"{Colors.GOTHAM_TEXT}{nome_campo.replace('_', ' ').capitalize()}: {Colors.ENDC}"
            valor_campo = input(input_prompt).strip()
            if not valor_campo and nome_campo in ['procedimento', 'procedimento01', 'procedimento02', 'data_disponivel',
                                                  'data_vigencia', 'cidade_estado']:
                print(
                    f"{Colors.WARNING_YELLOW}O campo '{nome_campo.replace('_', ' ')}' é obrigatório. Saindo.{Colors.ENDC}")
                return
            campos_a_preencher[placeholder] = valor_campo
    try:
        frase_gerada = frase_modelo
        for placeholder, valor in campos_a_preencher.items():
            if valor:
                frase_gerada = frase_gerada.replace(placeholder, valor)
        frase_limpa = os.linesep.join([s.strip() for s in frase_gerada.splitlines() if s.strip()])
        pyperclip.copy(frase_limpa)
        print(f"\n{Colors.SUCCESS_GREEN}✓ Fraseologia de Negativa copiada para a área de transferência!{Colors.ENDC}")
        print(f"\n{Colors.BOLD}{Colors.INFO_CYAN}----------------------------------------{Colors.ENDC}")
        print(f"{Colors.GOTHAM_TEXT}{frase_limpa}{Colors.ENDC}")
        print(f"{Colors.BOLD}{Colors.INFO_CYAN}----------------------------------------{Colors.ENDC}")
    except Exception as e:
        print(f"{Colors.ERROR_RED}Erro ao gerar fraseologia negativa: {e}{Colors.ENDC}")


# =================================================================================
# BLOCO 6: FUNÇÕES DE MENU E INTERFACE
# =================================================================================
# Este bloco contém as funções que controlam a exibição dos menus, a interação
# com o usuário e a navegação dentro do sistema.

def manter_sessao_fraseologia(initial_choice):
    """Permite ao usuário gerar múltiplas frases sem voltar ao menu principal."""
    while True:
        if initial_choice == 'F':
            gerar_fraseologia_positiva()
        elif initial_choice == 'N':
            gerar_fraseologia_negativa()
        elif initial_choice == 'R':
            gerar_texto_reembolso()
        elif initial_choice.upper() in ['V', 'VOLTAR']:
            break
        else:
            print(f"{Colors.ERROR_RED}✗ Erro: Opção inválida. Tente novamente.{Colors.ENDC}")

        print(f"\n{Colors.GOTHAM_TEXT}════════════════════════════════════════{Colors.ENDC}")
        print(
            f"{Colors.BOLD}{Colors.BATMAN_YELLOW}► [F]{Colors.ENDC} {Colors.GOTHAM_TEXT}Gerar outra Autorização{Colors.ENDC}")
        print(
            f"{Colors.BOLD}{Colors.BATMAN_YELLOW}► [N]{Colors.ENDC} {Colors.GOTHAM_TEXT}Gerar outra Negativa{Colors.ENDC}")
        print(
            f"{Colors.BOLD}{Colors.BATMAN_YELLOW}► [R]{Colors.ENDC} {Colors.GOTHAM_TEXT}Gerar outro Reembolso{Colors.ENDC}")
        print(
            f"{Colors.BOLD}{Colors.BATMAN_YELLOW}► [V]{Colors.ENDC} {Colors.GOTHAM_TEXT}Voltar ao menu principal{Colors.ENDC}")
        print(f"{Colors.GOTHAM_TEXT}════════════════════════════════════════{Colors.ENDC}")

        initial_choice = input(f"{Colors.GOTHAM_TEXT}Comando > {Colors.ENDC}").strip().upper()
        if initial_choice.upper() in ['V', 'VOLTAR']:
            break


def exibir_menu_ferramentas_texto(terminal_width):
    """Exibe o menu de ferramentas de texto e processa a escolha do usuário."""
    while True:
        os.system('cls' if os.name == 'nt' else 'clear')
        print(f"\n{Colors.GOTHAM_TEXT}{center_text('═' * 60, terminal_width)}{Colors.ENDC}")
        print(
            center_text(f"{Colors.BOLD}{Colors.INFO_CYAN}--- FERRAMENTAS DE TEXTO ---{Colors.ENDC}",
                        terminal_width + 12))
        print(
            f"\n{center_text(f'{Colors.UNDERLINE}{Colors.BATMAN_YELLOW}>>> OPÇÕES DISPONÍVEIS <<<{Colors.ENDC}', terminal_width + 2)}")
        print(
            f"  {Colors.BOLD}{Colors.BATMAN_YELLOW}► [F]{Colors.ENDC} {Colors.GOTHAM_TEXT}Gerar Fraseologia de Autorização{Colors.ENDC}")
        print(
            f"  {Colors.BOLD}{Colors.BATMAN_YELLOW}► [N]{Colors.ENDC} {Colors.GOTHAM_TEXT}Gerar Fraseologia de Negativa{Colors.ENDC}")
        print(
            f"  {Colors.BOLD}{Colors.BATMAN_YELLOW}► [R]{Colors.ENDC} {Colors.GOTHAM_TEXT}Orientação para Reembolso{Colors.ENDC}")
        print(
            f"\n  {Colors.BOLD}{Colors.BATMAN_YELLOW}► [V]{Colors.ENDC} {Colors.GOTHAM_TEXT}Voltar ao menu principal{Colors.ENDC}")
        print(f"{Colors.GOTHAM_TEXT}{center_text('═' * 60, terminal_width)}{Colors.ENDC}")

        escolha = input(f"{Colors.GOTHAM_TEXT}Comando > {Colors.ENDC}").strip().upper()
        if escolha == 'F':
            gerar_fraseologia_positiva()
        elif escolha == 'N':
            gerar_fraseologia_negativa()
        elif escolha == 'R':
            gerar_texto_reembolso()
        elif escolha in ['V', 'VOLTAR']:
            break
        else:
            print(f"{Colors.ERROR_RED}✗ Erro: Opção inválida. Tente novamente.{Colors.ENDC}")
        input(f"{Colors.GOTHAM_TEXT}Pressione ENTER para continuar...{Colors.ENDC}")


def exibir_menu_setores_dados(buscador, nomes_abas, terminal_width):
    """Exibe o menu para seleção de setores de dados e gerencia a busca."""
    while True:
        os.system('cls' if os.name == 'nt' else 'clear')
        print(f"\n{Colors.GOTHAM_TEXT}{center_text('═' * 60, terminal_width)}{Colors.ENDC}")
        print(
            center_text(f"{Colors.BOLD}{Colors.INFO_CYAN}--- SETORES DE DADOS ---{Colors.ENDC}",
                        terminal_width + 12))
        print(center_text(f"{Colors.WARNING_YELLOW}Selecione um setor de dados:{Colors.ENDC}",
                          terminal_width + 12))
        colunas = 2
        total_abas = len(nomes_abas)
        abas_por_coluna = (total_abas + colunas - 1) // colunas
        max_aba_length = max(len(aba) for aba in nomes_abas) if nomes_abas else 0
        col_width = max_aba_length + 8
        min_col_width = 30
        if col_width < min_col_width:
            col_width = min_col_width
        total_menu_width = colunas * (col_width + 2)
        left_margin = max(0, (terminal_width - total_menu_width) // 2)
        for i in range(abas_por_coluna):
            linha = " " * left_margin
            for j in range(colunas):
                index = i + j * abas_por_coluna
                if index < total_abas:
                    nome_aba = nomes_abas[index]
                    linha += f"  {Colors.BOLD}{Colors.BATMAN_YELLOW}► [{index + 1:02d}]{Colors.ENDC} {Colors.GOTHAM_TEXT}{nome_aba:<{col_width - 8}}{Colors.ENDC}"
                else:
                    linha += " " * (col_width + 2)
            if linha.strip():
                print(linha)
        print(
            f"\n  {Colors.BOLD}{Colors.BATMAN_YELLOW}► [V]{Colors.ENDC} {Colors.GOTHAM_TEXT}Voltar ao menu principal{Colors.ENDC}")
        print(f"{Colors.GOTHAM_TEXT}{center_text('═' * 60, terminal_width)}{Colors.ENDC}")

        escolha_aba = input(f"{Colors.GOTHAM_TEXT}Comando > {Colors.ENDC}").strip().upper()
        if escolha_aba in ['V', 'VOLTAR']:
            break
        if escolha_aba.isdigit() and 1 <= int(escolha_aba) <= len(nomes_abas):
            nome_aba_selecionada = nomes_abas[int(escolha_aba) - 1]
            print(f"\n{Colors.BOLD}{Colors.INFO_CYAN}--- STATUS DO SETOR ---{Colors.ENDC}")
            print(
                f"  {Colors.SUCCESS_GREEN}► Setor {Colors.BOLD}{nome_aba_selecionada}{Colors.ENDC} {Colors.SUCCESS_GREEN}ativado. Status: {Colors.BOLD}[ONLINE]{Colors.ENDC}")
            print(f"  {Colors.GOTHAM_TEXT}Aguardando consulta...{Colors.ENDC}")
            termos_voltar = ['V', 'VOLTAR']
            while True:
                termo = input(
                    f"\n{Colors.GOTHAM_TEXT}Termo de busca (digite '{' ou '.join(termos_voltar)}' para retornar ao menu):{Colors.ENDC}\n{Colors.GOTHAM_TEXT}➜ {Colors.ENDC}").strip()
                if termo.upper() in termos_voltar:
                    break
                if termo:
                    inicio = time.time()
                    resultados = buscador.buscar_avancada(nome_aba_selecionada, termo)
                    tempo = time.time() - inicio
                    buscador.exibir_resultados_avancados(resultados, termo, nome_aba_selecionada, tempo)
        else:
            print(f"{Colors.ERROR_RED}✗ Erro de sintaxe: Comando não reconhecido. Tente novamente.{Colors.ENDC}")
            input(f"{Colors.GOTHAM_TEXT}Pressione ENTER para continuar...{Colors.ENDC}")


def exibir_menu_ferramentas_sistema(buscador, terminal_width):
    """Exibe o menu de ferramentas do sistema e gerencia as ações."""
    while True:
        os.system('cls' if os.name == 'nt' else 'clear')
        print(f"\n{Colors.GOTHAM_TEXT}{center_text('═' * 60, terminal_width)}{Colors.ENDC}")
        print(
            center_text(f"{Colors.BOLD}{Colors.INFO_CYAN}--- FERRAMENTAS DO SISTEMA ---{Colors.ENDC}",
                        terminal_width + 12))
        print(
            f"\n{center_text(f'{Colors.UNDERLINE}{Colors.BATMAN_YELLOW}>>> OPÇÕES DISPONÍVEIS <<<{Colors.ENDC}', terminal_width + 2)}")
        print(
            f"  {Colors.BOLD}{Colors.BATMAN_YELLOW}► [ST]{Colors.ENDC} {Colors.GOTHAM_TEXT}Relatório de status{Colors.ENDC}")
        print(
            f"  {Colors.BOLD}{Colors.BATMAN_YELLOW}► [CA]{Colors.ENDC} {Colors.GOTHAM_TEXT}Limpar cache de busca{Colors.ENDC}")
        print(
            f"  {Colors.BOLD}{Colors.BATMAN_YELLOW}► [CFG]{Colors.ENDC} {Colors.GOTHAM_TEXT}Salvar configurações do sistema{Colors.ENDC}")
        print(
            f"  {Colors.BOLD}{Colors.BATMAN_YELLOW}► [RE]{Colors.ENDC} {Colors.GOTHAM_TEXT}Reiniciar sistema{Colors.ENDC}")
        print(
            f"\n  {Colors.BOLD}{Colors.BATMAN_YELLOW}► [V]{Colors.ENDC} {Colors.GOTHAM_TEXT}Voltar ao menu principal{Colors.ENDC}")
        print(f"{Colors.GOTHAM_TEXT}{center_text('═' * 60, terminal_width)}{Colors.ENDC}")

        escolha = input(f"{Colors.GOTHAM_TEXT}Comando > {Colors.ENDC}").strip().upper()
        if escolha == 'ST':
            buscador.mostrar_estatisticas()
        elif escolha == 'CA':
            buscador.limpar_cache()
        elif escolha == 'CFG':
            buscador.salvar_configuracao()
        elif escolha == 'RE':
            os.system('cls' if os.name == 'nt' else 'clear')
            print(f"{Colors.SUCCESS_GREEN}Reiniciando sistema...{Colors.ENDC}")
            main()
            return
        elif escolha in ['V', 'VOLTAR']:
            break
        else:
            print(f"{Colors.ERROR_RED}✗ Erro: Opção inválida. Tente novamente.{Colors.ENDC}")
        input(f"{Colors.GOTHAM_TEXT}Pressione ENTER para continuar...{Colors.ENDC}")


# =================================================================================
# BLOCO 7: PONTO DE ENTRADA DO PROGRAMA
# =================================================================================
# Este é o ponto de partida do script. A função 'main' é responsável por inicializar
# o sistema, carregar os dados e exibir o menu principal, controlando todo o fluxo.

def main():
    """Função principal que inicializa e executa o sistema."""
    warnings.filterwarnings('ignore')
    colorama.init()

    try:
        terminal_width = os.get_terminal_size().columns
    except OSError:
        terminal_width = 80

    print(f"\n{Colors.BOLD}{Colors.GOTHAM_TEXT}")
    print(center_text("INICIANDO MUNDO SOMBRIO - MKACETE", terminal_width))
    print(center_text("═" * terminal_width, terminal_width))
    print(center_text("SISTEMA SOMBRIO", terminal_width))
    print(f"{Colors.ENDC}")
    print(f"{Colors.GOTHAM_TEXT}Calibrando matriz de dados...{Colors.ENDC}")
    bar_length = terminal_width - 20
    for i in range(11):
        progress = i * 10
        filled_length = int(bar_length * i / 10)
        bar = "█" * filled_length + " " * (bar_length - filled_length)
        sys.stdout.write(f"\r{Colors.BATMAN_YELLOW}[{bar}]{Colors.ENDC} {Colors.BATMAN_YELLOW}{progress}%{Colors.ENDC}")
        sys.stdout.flush()
        time.sleep(0.1)
    sys.stdout.write(f"\r{Colors.SUCCESS_GREEN}Calibração concluída!{Colors.ENDC}\n")
    sys.stdout.flush()

    caminho_local = 'BATMAN.xlsx'
    caminho_completo = r'C:\Users\marcos.oliveira7\Documents\TREVAS\MKACETE-main\BATMAN.xlsx'

    if os.path.exists(caminho_local):
        nome_arquivo = caminho_local
        print(f"{Colors.SUCCESS_GREEN}✓ Conexão estabelecida com {Colors.BOLD}{caminho_local}{Colors.ENDC}")
    elif os.path.exists(caminho_completo):
        nome_arquivo = caminho_completo
        print(f"{Colors.SUCCESS_GREEN}✓ Conexão estabelecida com {Colors.BOLD}{caminho_completo}{Colors.ENDC}")
    else:
        print(f"{Colors.ERROR_RED}✗ ERRO FATAL: Matriz de dados 'BATMAN.xlsx' não encontrada.{Colors.ENDC}")
        print(f"  Verifique os caminhos: {caminho_local} ou {caminho_completo}")
        return

    buscador = MecanismoBuscaAvancado(nome_arquivo)
    if not buscador.dados_abas:
        print(f"{Colors.ERROR_RED}✗ ERRO: Falha ao carregar os setores de dados.{Colors.ENDC}")
        return

    nomes_abas = list(buscador.dados_abas.keys())

    while True:
        os.system('cls' if os.name == 'nt' else 'clear')
        print(f"\n{Colors.GOTHAM_TEXT}{center_text('═' * 60, terminal_width)}{Colors.ENDC}")
        print(
            center_text(f"{Colors.BOLD}{Colors.INFO_CYAN}--- TERMINAL SOMBRIO ---{Colors.ENDC}",
                        terminal_width + 12))
        print(center_text(f"{Colors.WARNING_YELLOW}Selecione uma página para continuar:{Colors.ENDC}",
                          terminal_width + 12))
        print(
            f"\n{center_text(f'{Colors.UNDERLINE}{Colors.BATMAN_YELLOW}>>> PÁGINAS <<<{Colors.ENDC}', terminal_width + 2)}")
        print(
            f"  {Colors.BOLD}{Colors.BATMAN_YELLOW}► [1]{Colors.ENDC} {Colors.GOTHAM_TEXT}Ferramentas de Texto{Colors.ENDC}")
        print(
            f"  {Colors.BOLD}{Colors.BATMAN_YELLOW}► [2]{Colors.ENDC} {Colors.GOTHAM_TEXT}Setores de Dados{Colors.ENDC}")
        print(
            f"  {Colors.BOLD}{Colors.BATMAN_YELLOW}► [3]{Colors.ENDC} {Colors.GOTHAM_TEXT}Ferramentas do Sistema{Colors.ENDC}")
        print(
            f"\n  {Colors.BOLD}{Colors.BATMAN_YELLOW}► [0]{Colors.ENDC} {Colors.GOTHAM_TEXT}Encerrar sistema{Colors.ENDC}")
        print(f"{Colors.GOTHAM_TEXT}{center_text('═' * 60, terminal_width)}{Colors.ENDC}")

        escolha = input(f"{Colors.GOTHAM_TEXT}Comando > {Colors.ENDC}").strip()

        if escolha == '0':
            print(
                f"\n{Colors.SUCCESS_GREEN}Sistema MKACETE encerrado. Volte sempre para mais alegria nos dados!{Colors.ENDC}")
            break
        elif escolha == '1':
            exibir_menu_ferramentas_texto(terminal_width)
        elif escolha == '2':
            exibir_menu_setores_dados(buscador, nomes_abas, terminal_width)
        elif escolha == '3':
            exibir_menu_ferramentas_sistema(buscador, terminal_width)
        else:
            print(f"{Colors.ERROR_RED}✗ Erro: Opção inválida. Tente novamente.{Colors.ENDC}")
            input(f"{Colors.GOTHAM_TEXT}Pressione ENTER para continuar...{Colors.ENDC}")


if __name__ == "__main__":
    main()