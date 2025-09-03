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

warnings.filterwarnings('ignore')

# Dicionário com todas as frases de negativa.
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

# Novo dicionário para as ferramentas de texto, incluindo a de Reembolso.
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


class Colors:
    """Define cores para estilizar o texto no terminal com tema hacker-like."""
    # Cores de base para o tema hacker
    GREEN_LIME = '\033[38;5;82m'  # Verde limão brilhante
    CYAN_LIGHT = '\033[38;5;45m'  # Ciano claro
    PURPLE_ACCENT = '\033[38;5;165m'  # Roxo suave para acentos
    YELLOW_WARN = '\033[38;5;220m'  # Amarelo para avisos
    RED_ERROR = '\033[38;5;196m'  # Vermelho vibrante para erros
    GRAY_TEXT = '\033[38;5;247m'  # Cinza claro para texto geral
    BRIGHT_MAGENTA = '\033[95m'  # Magenta brilhante
    BRIGHT_CYAN = '\033[96m'  # Ciano brilhante
    BRIGHT_GREEN = '\033[92m'  # Verde brilhante

    # Efeitos
    BOLD = '\033[1m'
    UNDERLINE = '\033[4m'
    ENDC = '\033[0m'  # Reset para as cores padrão do terminal


class MecanismoBuscaAvancado:
    """Gerencia o carregamento de dados e a busca inteligente em planilhas."""

    def __init__(self, nome_arquivo_excel):
        """Inicializa o buscador com o nome do arquivo Excel."""
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
        """Carrega configurações personalizadas do arquivo config.json."""
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
            print(f"{Colors.YELLOW_WARN}Aviso: Erro ao carregar configuração personalizada: {e}{Colors.ENDC}")

        return config_padrao

    def _carregar_dados(self):
        """Carrega e otimiza os dados do arquivo Excel."""
        if not os.path.exists(self.nome_arquivo_excel):
            print(
                f"{Colors.RED_ERROR}ERRO: Arquivo '{self.nome_arquivo_excel}' não detectado. Verifique a matriz de dados.{Colors.ENDC}")
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
            print(f"{Colors.GREEN_LIME}✓ Banco de dados online em {tempo_carregamento:.2f}s!{Colors.ENDC}")
            print(f"{Colors.CYAN_LIGHT}Total de abas carregadas: {len(self.dados_abas)}{Colors.ENDC}")


        except Exception as e:
            print(f"{Colors.RED_ERROR}Falha crítica ao carregar o arquivo: {e}{Colors.ENDC}")

    def _preparar_dados_busca(self):
        """
        Normaliza os dados para busca sem acentuação e mais rápida.
        Cria uma coluna unificada '_TEXTO_BUSCA' com todos os dados.
        """
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

    def _mostrar_estatisticas_abas(self):
        """Exibe o número de registros em cada aba carregada."""
        print(f"\n{Colors.UNDERLINE}Estatísticas de Setores:{Colors.ENDC}")
        for nome_aba, df in self.dados_abas.items():
            print(f"  {Colors.CYAN_LIGHT}{nome_aba}{Colors.ENDC}: {len(df)} registros")

    def buscar_avancada(self, nome_aba, termo, filtros=None, max_resultados=None):
        """Executa a busca com múltiplos algoritmos e filtros."""
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
        """Combina busca exata, por relevância e fuzzy para resultados abrangentes."""
        termo_limpo = unidecode(termo).lower().strip()

        resultados_exatos = self._busca_exata(df, termo_limpo)
        resultados_relevancia = self._busca_por_relevancia(df, termo_limpo, max_resultados)
        resultados_fuzzy = self._busca_fuzzy(df, termo_limpo, max_resultados)

        df_combinado = pd.concat([resultados_exatos, resultados_relevancia, resultados_fuzzy], axis=0)
        df_resultado = df_combinado.drop_duplicates(keep='first').head(max_resultados)

        return df_resultado

    def _busca_exata(self, df, termo):
        """Busca por correspondência exata do termo na coluna de busca unificada."""
        mascara = df['_TEXTO_BUSCA'].str.contains(termo, na=False)
        return df[mascara]

    def _busca_por_relevancia(self, df, termo, max_resultados):
        """Busca por relevância usando pesos em colunas prioritárias."""
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
        """Busca por similaridade de strings (fuzzy) em colunas prioritárias."""
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
        """Aplica filtros específicos ao DataFrame."""
        df_filtrado = df.copy()

        for coluna, valor in filtros.items():
            if coluna in df_filtrado.columns and valor:
                mascara = df_filtrado[coluna].str.contains(str(valor), na=False)
                df_filtrado = df_filtrado[mascara]

        return df_filtrado

    def exibir_resultados_avancados(self, resultados, termo, nome_aba, tempo_busca=None):
        """Exibe resultados com formatação detalhada e estatísticas."""
        if resultados.empty:
            print(
                f"\n{Colors.YELLOW_WARN}Nenhum resultado encontrado para '{termo}' no setor '{nome_aba}'.{Colors.ENDC}")
            return

        print(f"\n{Colors.CYAN_LIGHT}>>> Resultado da Busca Terminal <<<{Colors.ENDC}")
        print(f"{Colors.GREEN_LIME}┌─ Termo de busca: {Colors.BOLD}{termo}{Colors.ENDC}")
        print(f"{Colors.GREEN_LIME}├─ Setor de dados: {Colors.BOLD}{nome_aba}{Colors.ENDC}")
        print(f"{Colors.GREEN_LIME}├─ Total de resultados: {Colors.BOLD}{len(resultados)}{Colors.ENDC}")
        if tempo_busca:
            print(f"{Colors.GREEN_LIME}└─ Tempo de execução: {Colors.BOLD}{tempo_busca:.3f}s{Colors.ENDC}")
        print(f"{Colors.CYAN_LIGHT}-----------------------------------------{Colors.ENDC}")

        if 'infiltracao' in nome_aba.lower() or 'medicacao' in nome_aba.lower():
            self._exibir_resultados_por_tema(resultados)
        else:
            self._exibir_generico_futurista(resultados)

        print(f"\n{Colors.CYAN_LIGHT}-----------------------------------------{Colors.ENDC}")

    def _exibir_resultados_por_tema(self, resultados):
        """
        Organiza e exibe os resultados com base em temas de busca.
        Mais flexível para diferentes abas de infiltração/medicação.
        """
        temas = {
            "INFILTRAÇÃO / BLOQUEIO": ['infiltracao', 'infiltração', 'bloqueio'],
            "RETIRADA DE MEDICAMENTO": ['retirada', 'remocao', 'remover'],
        }

        # Cria uma cópia para evitar modificar o DataFrame original no loop
        df_restantes = resultados.copy()

        for titulo_tema, palavras_chave in temas.items():
            # Cria uma máscara para encontrar as linhas que contêm as palavras-chave do tema
            mask = df_restantes['_TEXTO_BUSCA'].str.contains('|'.join(palavras_chave), na=False)
            df_tema = df_restantes[mask]

            if not df_tema.empty:
                self._exibir_bloco_categoria(titulo_tema, df_tema)

            # Remove os resultados já exibidos do DataFrame de 'restantes'
            df_restantes = df_restantes[~mask]

        # Exibe quaisquer resultados que não se encaixaram em nenhum tema
        if not df_restantes.empty:
            self._exibir_bloco_categoria("Outros Resultados (Sem Categoria)", df_restantes)

    def _exibir_bloco_categoria(self, titulo, df):
        """
        Exibe um bloco de resultados com um título e formatação padronizada.
        """
        if not df.empty:
            print(f"\n{Colors.BOLD}{Colors.PURPLE_ACCENT}----- {titulo} ({len(df)}) -----{Colors.ENDC}")
            self._exibir_generico_futurista(df)

    def _exibir_generico_futurista(self, df):
        """Exibe os dados de forma futurista, com cores e organização."""
        for i, (_, row) in enumerate(df.iterrows(), 1):
            print(
                f"\n{Colors.BOLD}{Colors.CYAN_LIGHT}═══════════════ [{i:03d}] Resultado ═══════════════{Colors.ENDC}")

            key_color = Colors.GRAY_TEXT
            value_color = Colors.GREEN_LIME

            # Colunas prioritárias para as abas 17 e 18
            colunas_prioritarias = ['PRESTADOR', 'ZONA/REGIÃO', 'CNPJ', 'CD PESSOA']

            # Exibe as colunas prioritárias primeiro
            for col in colunas_prioritarias:
                if col in row and str(row[col]).strip():
                    print(f"  {key_color}{str(col).upper()}:{Colors.ENDC} {value_color}{row[col]}{Colors.ENDC}")

            # Exibe as demais colunas
            for col, val in row.items():
                if str(val).strip() and col != '_TEXTO_BUSCA' and col not in colunas_prioritarias:
                    print(f"  {key_color}{str(col).upper()}:{Colors.ENDC} {value_color}{val}{Colors.ENDC}")

            print(
                f"{Colors.BOLD}{Colors.CYAN_LIGHT}═══════════════════════════════════════════════════════{Colors.ENDC}")

    def mostrar_estatisticas(self):
        """Mostra estatísticas de uso do sistema."""
        print(f"\n{Colors.CYAN_LIGHT}>>> Relatório de Status do Sistema <<<{Colors.ENDC}")
        print(
            f"{Colors.GRAY_TEXT}Total de consultas: {Colors.BOLD}{self.estatisticas['total_buscas']}{Colors.ENDC}")
        print(
            f"{Colors.GRAY_TEXT}Tempo médio de resposta: {Colors.BOLD}{self.estatisticas['tempo_medio_busca']:.3f}s{Colors.ENDC}")
        print(
            f"{Colors.GRAY_TEXT}Total de resultados gerados: {Colors.BOLD}{self.estatisticas['resultados_encontrados']}{Colors.ENDC}")
        print(f"{Colors.GRAY_TEXT}Cache de memória: {Colors.BOLD}{len(self.cache_busca)}{Colors.ENDC}")
        print(
            f"{Colors.GRAY_TEXT}Status do cache: {Colors.BOLD}{'Ativo' if self.config['habilitar_cache'] else 'Inativo'}{Colors.ENDC}")
        print(f"{Colors.CYAN_LIGHT}-----------------------------------------{Colors.ENDC}")

    def limpar_cache(self):
        """Limpa o cache de busca."""
        self.cache_busca.clear()
        print(f"{Colors.GREEN_LIME}✓ Cache de busca limpo com sucesso! {Colors.ENDC}")

    def salvar_configuracao(self):
        """Salva a configuração atual em arquivo JSON."""
        try:
            with open('config.json', 'w', encoding='utf-8') as f:
                json.dump(self.config, f, indent=2, ensure_ascii=False)
            print(f"{Colors.GREEN_LIME}✓ Configuração do sistema salva em 'config.json'!{Colors.ENDC}")
        except Exception as e:
            print(f"{Colors.RED_ERROR}Falha ao salvar a configuração: {e}{Colors.ENDC}")


def gerar_texto_reembolso():
    """Gera e copia o texto de reembolso."""
    try:
        frase_reembolso = FERRAMENTAS_TEXTO['REEMBOLSO']['fraseologia']
        pyperclip.copy(frase_reembolso)
        print(f"\n{Colors.GREEN_LIME}✓ Texto de Reembolso copiado para a área de transferência!{Colors.ENDC}")
        print(f"\n{Colors.BOLD}{Colors.CYAN_LIGHT}----------------------------------------{Colors.ENDC}")
        print(f"{Colors.GRAY_TEXT}{frase_reembolso}{Colors.ENDC}")
        print(f"{Colors.BOLD}{Colors.CYAN_LIGHT}----------------------------------------{Colors.ENDC}")
    except Exception as e:
        print(f"{Colors.RED_ERROR}Erro ao gerar texto de reembolso: {e}{Colors.ENDC}")


def gerar_fraseologia_positiva():
    """Gera e copia uma fraseologia positiva para um ou mais procedimentos."""
    try:
        try:
            num_procedimentos_str = input(
                f"{Colors.CYAN_LIGHT}Quantos procedimentos serão autorizados? {Colors.ENDC}").strip()
            num_procedimentos = int(num_procedimentos_str)
            if num_procedimentos <= 0:
                print(f"{Colors.YELLOW_WARN}O número de procedimentos deve ser maior que zero. Saindo.{Colors.ENDC}")
                return
        except ValueError:
            print(f"{Colors.RED_ERROR}Entrada inválida. Digite um número inteiro.{Colors.ENDC}")
            return

        fraseologia_base = """
PREZADO(A) SR(A). [_NM_BENEFICIARIO_],
NÚMERO DO PROTOCOLO: [_NU_PROTOCOLO_]

SUA SOLICITAÇÃO DE AUTORIZAÇÃO PARA EXAME FOI RECEBIDA COM OS SEGUINTES DADOS:
"""
        corpo_fraseologia = ""
        for i in range(num_procedimentos):
            procedimento = input(
                f"{Colors.CYAN_LIGHT}Procedimento ({i + 1}/{num_procedimentos}): {Colors.ENDC}").strip().upper()
            senha = input(f"{Colors.CYAN_LIGHT}Senha ({i + 1}/{num_procedimentos}): {Colors.ENDC}").strip()
            prestador = input(
                f"{Colors.CYAN_LIGHT}Prestador ({i + 1}/{num_procedimentos}): {Colors.ENDC}").strip().upper()

            if not procedimento or not senha or not prestador:
                print(f"{Colors.YELLOW_WARN}Todos os campos são obrigatórios. Saindo.{Colors.ENDC}")
                return

            corpo_fraseologia += f"\nPROCEDIMENTO: {procedimento}"
            corpo_fraseologia += f"\nSENHA: {senha}"
            corpo_fraseologia += f"\nPRESTADOR: {prestador}"

            if i < num_procedimentos - 1:
                corpo_fraseologia += "\n================="

        frase_final = fraseologia_base + corpo_fraseologia
        frase_limpa = os.linesep.join([s.strip() for s in frase_final.splitlines() if s.strip()])
        pyperclip.copy(frase_limpa)

        print(f"\n{Colors.GREEN_LIME}✓ Fraseologia de Autorização copiada para a área de transferência!{Colors.ENDC}")
        print(f"\n{Colors.BOLD}{Colors.CYAN_LIGHT}----------------------------------------{Colors.ENDC}")
        print(f"{Colors.GRAY_TEXT}{frase_limpa}{Colors.ENDC}")
        print(f"{Colors.BOLD}{Colors.CYAN_LIGHT}----------------------------------------{Colors.ENDC}")

    except Exception as e:
        print(f"{Colors.RED_ERROR}Erro ao gerar fraseologia positiva: {e}{Colors.ENDC}")


def gerar_fraseologia_negativa():
    """Gera e copia uma fraseologia de negativa."""
    print(f"\n{Colors.BOLD}{Colors.CYAN_LIGHT}--- Ferramenta de Negativa ---{Colors.ENDC}")
    print(f"{Colors.GRAY_TEXT}Códigos de negativa disponíveis:{Colors.ENDC}")
    for codigo, info in DADOS_RESTRICOES.items():
        print(f"{Colors.YELLOW_WARN}[{codigo}]{Colors.ENDC} - {info['nome']}")

    escolha_negativa = input(f"{Colors.CYAN_LIGHT}Escolha o código de negativa: {Colors.ENDC}").strip()
    if escolha_negativa not in DADOS_RESTRICOES:
        print(f"{Colors.RED_ERROR}Código de negativa inválido.{Colors.ENDC}")
        return

    info_negativa = DADOS_RESTRICOES[escolha_negativa]
    frase_modelo = info_negativa['fraseologia']

    campos_a_preencher = {}

    # Identifica os campos necessários na fraseologia
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
                input_prompt = f"{Colors.CYAN_LIGHT}{nome_campo.replace('_', ' ').capitalize()} (dd/mm/aaaa): {Colors.ENDC}"
            elif nome_campo == 'cidade_estado':
                input_prompt = f"{Colors.CYAN_LIGHT}Cidade ou Estado: {Colors.ENDC}"
            else:
                input_prompt = f"{Colors.CYAN_LIGHT}{nome_campo.replace('_', ' ').capitalize()}: {Colors.ENDC}"

            valor_campo = input(input_prompt).strip()

            # Validação de campos obrigatórios
            if not valor_campo and nome_campo in ['procedimento', 'procedimento01', 'procedimento02', 'data_disponivel',
                                                  'data_vigencia', 'cidade_estado']:
                print(
                    f"{Colors.YELLOW_WARN}O campo '{nome_campo.replace('_', ' ')}' é obrigatório. Saindo.{Colors.ENDC}")
                return
            campos_a_preencher[placeholder] = valor_campo

    try:
        frase_gerada = frase_modelo
        for placeholder, valor in campos_a_preencher.items():
            if valor:
                frase_gerada = frase_gerada.replace(placeholder, valor)

        frase_limpa = os.linesep.join([s.strip() for s in frase_gerada.splitlines() if s.strip()])
        pyperclip.copy(frase_limpa)
        print(f"\n{Colors.GREEN_LIME}✓ Fraseologia de Negativa copiada para a área de transferência!{Colors.ENDC}")
        print(f"\n{Colors.BOLD}{Colors.CYAN_LIGHT}----------------------------------------{Colors.ENDC}")
        print(f"{Colors.GRAY_TEXT}{frase_limpa}{Colors.ENDC}")
        print(f"{Colors.BOLD}{Colors.CYAN_LIGHT}----------------------------------------{Colors.ENDC}")

    except Exception as e:
        print(f"{Colors.RED_ERROR}Erro ao gerar fraseologia negativa: {e}{Colors.ENDC}")


def manter_sessao_fraseologia(initial_choice):
    """Mantém a sessão de geração de fraseologia."""
    while True:
        # Executa a primeira escolha do usuário
        if initial_choice == 'F':
            gerar_fraseologia_positiva()
        elif initial_choice == 'N':
            gerar_fraseologia_negativa()
        elif initial_choice == 'R':
            gerar_texto_reembolso()
        elif initial_choice.upper() in ['V', 'VOLTAR']:
            break
        else:
            print(f"{Colors.RED_ERROR}✗ Erro: Opção inválida. Tente novamente.{Colors.ENDC}")

        print(f"\n{Colors.GRAY_TEXT}════════════════════════════════════════{Colors.ENDC}")
        print(f"{Colors.BRIGHT_MAGENTA}► [F]{Colors.ENDC} Gerar outra Autorização")
        print(f"{Colors.BRIGHT_MAGENTA}► [N]{Colors.ENDC} Gerar outra Negativa")
        print(f"{Colors.BRIGHT_MAGENTA}► [R]{Colors.ENDC} Gerar outro Reembolso")
        print(f"{Colors.BRIGHT_MAGENTA}► [V]{Colors.ENDC} Voltar ao menu principal")
        print(f"{Colors.GRAY_TEXT}════════════════════════════════════════{Colors.ENDC}")

        initial_choice = input(f"{Colors.GREEN_LIME}Comando > {Colors.ENDC}").strip().upper()
        if initial_choice.upper() in ['V', 'VOLTAR']:
            break


def main():
    """Função principal que executa a interface de linha de comando."""
    colorama.init()

    try:
        terminal_width = os.get_terminal_size().columns
    except OSError:
        terminal_width = 80

    def center_text(text, width):
        return text.center(width)

    print(f"\n{Colors.BOLD}{Colors.GREEN_LIME}")
    print(center_text("INICIANDO MUNDO SOMBRIO - MKACETE", terminal_width))
    print(center_text("═" * terminal_width, terminal_width))
    print(center_text("SISTEMA SOMBRIO", terminal_width))
    print(f"{Colors.ENDC}")

    # Animação de carregamento
    print(f"{Colors.CYAN_LIGHT}Calibrando matriz de dados...{Colors.ENDC}")
    bar_length = terminal_width - 20
    for i in range(11):
        progress = i * 10
        filled_length = int(bar_length * i / 10)
        bar = "█" * filled_length + " " * (bar_length - filled_length)
        sys.stdout.write(f"\r{Colors.GREEN_LIME}[{bar}]{Colors.ENDC} {progress}%")
        sys.stdout.flush()
        time.sleep(0.2)
    sys.stdout.write(f"\r{Colors.GREEN_LIME}Calibração concluída!{Colors.ENDC}\n")
    sys.stdout.flush()

    caminho_local = 'BATMAN.xlsx'
    caminho_completo = r'C:\Users\marcos.oliveira7\Documents\TREVAS\MKACETE-main\BATMAN.xlsx'

    if os.path.exists(caminho_local):
        nome_arquivo = caminho_local
        print(f"{Colors.GREEN_LIME}✓ Conexão estabelecida com {caminho_local}{Colors.ENDC}")
    elif os.path.exists(caminho_completo):
        nome_arquivo = caminho_completo
        print(f"{Colors.GREEN_LIME}✓ Conexão estabelecida com {caminho_completo}{Colors.ENDC}")
    else:
        print(f"{Colors.RED_ERROR}✗ ERRO FATAL: Matriz de dados 'BATMAN.xlsx' não encontrada.{Colors.ENDC}")
        print(f"  Verifique os caminhos: {caminho_local} ou {caminho_completo}")
        return

    buscador = MecanismoBuscaAvancado(nome_arquivo)

    if not buscador.dados_abas:
        print(f"{Colors.RED_ERROR}✗ ERRO: Falha ao carregar os setores de dados.{Colors.ENDC}")
        return

    nomes_abas = list(buscador.dados_abas.keys())

    while True:
        print(f"\n{Colors.GRAY_TEXT}{center_text('═' * 60, terminal_width)}{Colors.ENDC}")
        print(
            center_text(f"{Colors.BOLD}{Colors.CYAN_LIGHT}--- TERMINAL MKACETE v1.0 ---{Colors.ENDC}",
                        terminal_width + 12))
        print(center_text(f"{Colors.YELLOW_WARN}Selecione uma opção ou setor de dados:{Colors.ENDC}",
                          terminal_width + 12))

        # Agrupar e exibir as opções de Fraseologia
        print(
            f"\n{center_text(f'{Colors.UNDERLINE}{Colors.PURPLE_ACCENT}>>> FERRAMENTAS DE TEXTO <<<{Colors.ENDC}', terminal_width + 2)}")
        print(f"  {Colors.BRIGHT_MAGENTA}► [F]{Colors.ENDC} Gerar Fraseologia de Autorização")
        print(f"  {Colors.BRIGHT_MAGENTA}► [N]{Colors.ENDC} Gerar Fraseologia de Negativa")
        print(f"  {Colors.BRIGHT_MAGENTA}► [R]{Colors.ENDC} Orientação para Reembolso")

        # Exibir as abas de dados em colunas
        print(
            f"\n{center_text(f'{Colors.UNDERLINE}{Colors.PURPLE_ACCENT}>>> SETORES DE DADOS <<<{Colors.ENDC}', terminal_width + 2)}")
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
                    linha += f"  {Colors.BOLD}{Colors.BRIGHT_CYAN}► [{index + 1:02d}]{Colors.ENDC} {Colors.GRAY_TEXT}{nome_aba:<{col_width - 8}}{Colors.ENDC}"
                else:
                    linha += " " * (col_width + 2)

            if linha.strip():
                print(linha)

        # Agrupar e exibir as opções de funções do sistema
        print(
            f"\n{center_text(f'{Colors.UNDERLINE}{Colors.PURPLE_ACCENT}>>> FERRAMENTAS DO SISTEMA <<<{Colors.ENDC}', terminal_width + 2)}")
        print(f"  {Colors.BRIGHT_MAGENTA}► [S]{Colors.ENDC} Relatório de status")
        print(f"  {Colors.BRIGHT_MAGENTA}► [C]{Colors.ENDC} Limpar cache de busca")
        print(f"  {Colors.BRIGHT_MAGENTA}► [CFG]{Colors.ENDC} Salvar configurações do sistema")
        print(f"  {Colors.BRIGHT_MAGENTA}► [0]{Colors.ENDC} Encerrar sistema")

        print(f"{Colors.GRAY_TEXT}{center_text('═' * 60, terminal_width)}{Colors.ENDC}")

        escolha = input(f"{Colors.GREEN_LIME}Comando > {Colors.ENDC}").strip().upper()

        if escolha == '0':
            print(
                f"\n{Colors.GREEN_LIME}Sistema MKACETE encerrado. Volte sempre para mais alegria nos dados!{Colors.ENDC}")
            break
        elif escolha == 'S':
            buscador.mostrar_estatisticas()
        elif escolha == 'C':
            buscador.limpar_cache()
        elif escolha == 'CFG':
            buscador.salvar_configuracao()
        elif escolha in ['F', 'N', 'R']:
            manter_sessao_fraseologia(escolha)
        elif escolha.isdigit() and 1 <= int(escolha) <= len(nomes_abas):
            nome_aba_selecionada = nomes_abas[int(escolha) - 1]
            print(f"\n{Colors.BOLD}{Colors.CYAN_LIGHT}--- STATUS DO SETOR ---{Colors.ENDC}")
            print(
                f"  {Colors.GREEN_LIME}► Setor {Colors.BOLD}{nome_aba_selecionada}{Colors.ENDC} {Colors.GREEN_LIME}ativado. Status: {Colors.BOLD}[ONLINE]{Colors.ENDC}")
            print(f"  {Colors.GRAY_TEXT}Aguardando consulta...{Colors.ENDC}")

            termos_voltar = ['V', 'VOLTAR']

            while True:
                termo = input(
                    f"\n{Colors.CYAN_LIGHT}Termo de busca (digite '{' ou '.join(termos_voltar)}' para retornar ao menu):{Colors.ENDC}\n{Colors.GREEN_LIME}➜ {Colors.ENDC}").strip()
                if termo.upper() in termos_voltar:
                    break
                if termo:
                    inicio = time.time()
                    resultados = buscador.buscar_avancada(nome_aba_selecionada, termo)
                    tempo = time.time() - inicio
                    buscador.exibir_resultados_avancados(resultados, termo, nome_aba_selecionada, tempo)
        else:
            print(f"{Colors.RED_ERROR}✗ Erro de sintaxe: Comando não reconhecido. Tente novamente.{Colors.ENDC}")


if __name__ == "__main__":
    main()