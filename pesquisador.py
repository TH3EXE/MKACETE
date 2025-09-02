import pandas as pd
import os
import time
from unidecode import unidecode
from difflib import SequenceMatcher
import json
import warnings
import colorama

warnings.filterwarnings('ignore')


class Colors:
    """Define cores para estilizar o texto no terminal."""
    HEADER = '\033[95m'
    OKBLUE = '\033[94m'
    OKCYAN = '\033[96m'
    OKGREEN = '\033[92m'
    WARNING = '\033[93m'
    FAIL = '\033[91m'
    ENDC = '\033[0m'
    BOLD = '\033[1m'
    UNDERLINE = '\033[4m'
    LIGHT_CYAN = '\033[96m'
    LIGHT_GREEN = '\033[92m'
    LIGHT_BLUE = '\033[94m'
    PURPLE = '\033[95m'
    YELLOW = '\033[93m'
    RED = '\033[91m'


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
            print(f"{Colors.WARNING}Aviso: Erro ao carregar configuração personalizada: {e}{Colors.ENDC}")

        return config_padrao

    def _carregar_dados(self):
        """Carrega e otimiza os dados do arquivo Excel."""
        print(f"{Colors.LIGHT_CYAN}Iniciando a calibração do banco de dados...{Colors.ENDC}")

        if not os.path.exists(self.nome_arquivo_excel):
            print(
                f"{Colors.RED}ERRO: Arquivo '{self.nome_arquivo_excel}' não detectado. Verifique a matriz de dados.{Colors.ENDC}")
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
            print(f"{Colors.LIGHT_GREEN}✓ Banco de dados online em {tempo_carregamento:.2f}s!{Colors.ENDC}")
            print(f"{Colors.LIGHT_BLUE}Total de abas carregadas: {len(self.dados_abas)}{Colors.ENDC}")

            self._mostrar_estatisticas_abas()

        except Exception as e:
            print(f"{Colors.RED}Falha crítica ao carregar o arquivo: {e}{Colors.ENDC}")

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
            print(f"  {Colors.OKCYAN}{nome_aba}{Colors.ENDC}: {len(df)} registros")

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
            print(f"\n{Colors.YELLOW}Nenhum resultado encontrado para '{termo}' no setor '{nome_aba}'.{Colors.ENDC}")
            return

        print(f"\n{Colors.LIGHT_CYAN}>>> Resultado da Busca Terminal <<<{Colors.ENDC}")
        print(f"{Colors.LIGHT_GREEN}┌─ Termo de busca: {Colors.BOLD}{termo}{Colors.ENDC}")
        print(f"{Colors.LIGHT_GREEN}├─ Setor de dados: {Colors.BOLD}{nome_aba}{Colors.ENDC}")
        print(f"{Colors.LIGHT_GREEN}├─ Total de resultados: {Colors.BOLD}{len(resultados)}{Colors.ENDC}")
        if tempo_busca:
            print(f"{Colors.LIGHT_GREEN}└─ Tempo de execução: {Colors.BOLD}{tempo_busca:.3f}s{Colors.ENDC}")
        print(f"{Colors.LIGHT_CYAN}-----------------------------------------{Colors.ENDC}")

        if 'infiltracao' in nome_aba.lower() or 'medicacao' in nome_aba.lower():
            self._exibir_resultados_por_tema(resultados)
        else:
            self._exibir_generico_futurista(resultados)

        print(f"\n{Colors.LIGHT_CYAN}-----------------------------------------{Colors.ENDC}")

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
            print(f"\n{Colors.BOLD}{Colors.PURPLE}----- {titulo} ({len(df)}) -----{Colors.ENDC}")
            self._exibir_generico_futurista(df)

    def _exibir_generico_futurista(self, df):
        """Exibe os dados de forma futurista, com cores e organização."""
        for i, (_, row) in enumerate(df.iterrows(), 1):
            print(
                f"\n{Colors.BOLD}{Colors.LIGHT_CYAN}═══════════════ [{i:03d}] Resultado ═══════════════{Colors.ENDC}")

            key_color = Colors.LIGHT_BLUE
            value_color = Colors.OKGREEN

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
                f"{Colors.BOLD}{Colors.LIGHT_CYAN}═══════════════════════════════════════════════════════{Colors.ENDC}")

    def mostrar_estatisticas(self):
        """Mostra estatísticas de uso do sistema."""
        print(f"\n{Colors.LIGHT_CYAN}>>> Relatório de Status do Sistema <<<{Colors.ENDC}")
        print(
            f"{Colors.LIGHT_BLUE}Total de consultas: {Colors.BOLD}{self.estatisticas['total_buscas']}{Colors.ENDC}")
        print(
            f"{Colors.LIGHT_BLUE}Tempo médio de resposta: {Colors.BOLD}{self.estatisticas['tempo_medio_busca']:.3f}s{Colors.ENDC}")
        print(
            f"{Colors.LIGHT_BLUE}Total de resultados gerados: {Colors.BOLD}{self.estatisticas['resultados_encontrados']}{Colors.ENDC}")
        print(f"{Colors.LIGHT_BLUE}Cache de memória: {Colors.BOLD}{len(self.cache_busca)}{Colors.ENDC}")
        print(
            f"{Colors.LIGHT_BLUE}Status do cache: {Colors.BOLD}{'Ativo' if self.config['habilitar_cache'] else 'Inativo'}{Colors.ENDC}")
        print(f"{Colors.LIGHT_CYAN}-----------------------------------------{Colors.ENDC}")

    def limpar_cache(self):
        """Limpa o cache de busca."""
        self.cache_busca.clear()
        print(f"{Colors.LIGHT_GREEN}✓ Cache de busca limpo com sucesso! {Colors.ENDC}")

    def salvar_configuracao(self):
        """Salva a configuração atual em arquivo JSON."""
        try:
            with open('config.json', 'w', encoding='utf-8') as f:
                json.dump(self.config, f, indent=2, ensure_ascii=False)
            print(f"{Colors.LIGHT_GREEN}✓ Configuração do sistema salva em 'config.json'!{Colors.ENDC}")
        except Exception as e:
            print(f"{Colors.RED}Falha ao salvar a configuração: {e}{Colors.ENDC}")


def gerar_fraseologia():
    """Gera uma fraseologia com base em informações fornecidas pelo usuário."""
    print(f"\n{Colors.LIGHT_CYAN}--- GERADOR DE FRASEOLOGIA ---{Colors.ENDC}")
    print(f"{Colors.YELLOW}Preencha os campos para criar a fraseologia.{Colors.ENDC}")

    # Coleta as informações iniciais
    try:
        num_procedimentos = int(
            input(f"{Colors.LIGHT_BLUE}Quantos procedimentos deseja adicionar? {Colors.ENDC}").strip())
        if num_procedimentos <= 0:
            print(f"{Colors.RED}O número de procedimentos deve ser maior que zero.{Colors.ENDC}")
            return
    except ValueError:
        print(f"{Colors.RED}Entrada inválida. Por favor, digite um número inteiro.{Colors.ENDC}")
        return

    # Inicia a construção da fraseologia
    fraseologia_completa = f"""
Prezado(a) Sr(a). [_NM_BENEFICIARIO_]
NÚMERO DO PROTOCOLO: [_NU_PROTOCOLO_]

SUA SOLICITAÇÃO DE AUTORIZAÇÃO PARA EXAME FOI RECEBIDA COM OS SEGUINTES DADOS:
"""

    # Coleta os dados para cada procedimento e os adiciona à fraseologia
    for i in range(num_procedimentos):
        print(f"\n{Colors.LIGHT_BLUE}--- Dados do Procedimento {i + 1} ---{Colors.ENDC}")
        procedimento = input(f"{Colors.OKCYAN}Procedimento: {Colors.ENDC}").strip()
        senha = input(f"{Colors.OKCYAN}Senha: {Colors.ENDC}").strip()
        prestador = input(f"{Colors.OKCYAN}Prestador: {Colors.ENDC}").strip()

        fraseologia_completa += f"""
PROCEDIMENTO: {procedimento}
SENHA: {senha}
PRESTADOR: {prestador}
"""

    fraseologia_completa += f"""
EM CASO DE DÚVIDAS, POR FAVOR, ENTRE EM CONTATO COM A CENTRAL DE ATENDIMENTO PELOS TELEFONES: 4090-1740, 0800 409 1740 OU 0800 463 4648.
"""

    print(f"\n{Colors.LIGHT_GREEN}✓ Fraseologia gerada com sucesso!{Colors.ENDC}")
    print(f"{Colors.LIGHT_CYAN}--- COPIE E COLE ---{Colors.ENDC}")
    print(fraseologia_completa)
    print(f"{Colors.LIGHT_CYAN}--------------------{Colors.ENDC}")
    input(f"\n{Colors.OKBLUE}Pressione ENTER para voltar ao menu principal...{Colors.ENDC}")


def main():
    """Função principal que executa a interface de linha de comando."""
    colorama.init()

    try:
        terminal_width = os.get_terminal_size().columns
    except OSError:
        terminal_width = 80

    def center_text(text, width):
        return f"{text}".center(width)

    print(f"\n{Colors.BOLD}{Colors.LIGHT_BLUE}")
    print(center_text("═" * 60, terminal_width))
    print(center_text("MKACETE: SUA ALEGRIA NA ANÁLISE DE DADOS!", terminal_width))
    print(center_text("═" * 60, terminal_width))
    print(f"{Colors.ENDC}")

    caminho_local = 'BATMAN.xlsx'
    caminho_completo = r'C:\Users\Administrador\Documents\BATMAN.xlsx'

    if os.path.exists(caminho_local):
        nome_arquivo = caminho_local
        print(f"{Colors.LIGHT_GREEN}✓ Conexão estabelecida com {caminho_local}{Colors.ENDC}")
    elif os.path.exists(caminho_completo):
        nome_arquivo = caminho_completo
        print(f"{Colors.LIGHT_GREEN}✓ Conexão estabelecida com {caminho_completo}{Colors.ENDC}")
    else:
        print(f"{Colors.RED}✗ ERRO FATAL: Matriz de dados 'BATMAN.xlsx' não encontrada.{Colors.ENDC}")
        print(f"  Verifique os caminhos: {caminho_local} ou {caminho_completo}")
        return

    buscador = MecanismoBuscaAvancado(nome_arquivo)

    if not buscador.dados_abas:
        print(f"{Colors.RED}✗ ERRO: Falha ao carregar os setores de dados.{Colors.ENDC}")
        return

    nomes_abas = list(buscador.dados_abas.keys())

    while True:
        print(f"\n{Colors.LIGHT_BLUE}{center_text('═' * 60, terminal_width)}{Colors.ENDC}")
        print(
            center_text(f"{Colors.BOLD}{Colors.LIGHT_CYAN}--- MENU DO TERMINAL ---{Colors.ENDC}", terminal_width + 12))
        print(center_text(f"{Colors.YELLOW}Selecione uma opção ou setor de dados:{Colors.ENDC}",
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
                    linha += f"  {Colors.OKCYAN}► [{index + 1:02d}]{Colors.ENDC} {nome_aba:<{col_width - 8}}"
                else:
                    linha += " " * (col_width + 2)

            if linha.strip():
                print(linha)

        # Adiciona a nova opção no menu
        print(center_text(f"\n  {Colors.OKCYAN}► [F]{Colors.ENDC} Gerar Fraseologia", terminal_width + 12))
        print(center_text(f"  {Colors.OKCYAN}► [S]{Colors.ENDC} Relatório de status", terminal_width + 12))
        print(center_text(f"  {Colors.OKCYAN}► [C]{Colors.ENDC} Limpar cache de busca", terminal_width + 12))
        print(
            center_text(f"  {Colors.OKCYAN}► [CFG]{Colors.ENDC} Salvar configurações do sistema", terminal_width + 12))
        print(center_text(f"  {Colors.OKCYAN}► [0]{Colors.ENDC} Encerrar sistema", terminal_width + 12))
        print(f"{Colors.LIGHT_BLUE}{center_text('═' * 60, terminal_width)}{Colors.ENDC}")

        escolha = input(f"{Colors.LIGHT_GREEN}Comando: {Colors.ENDC}").strip().upper()

        if escolha == '0':
            print(
                f"\n{Colors.LIGHT_GREEN}Sistema MKACETE encerrado. Volte sempre para mais alegria nos dados!{Colors.ENDC}")
            break
        elif escolha == 'S':
            buscador.mostrar_estatisticas()
        elif escolha == 'C':
            buscador.limpar_cache()
        elif escolha == 'CFG':
            buscador.salvar_configuracao()
        elif escolha == 'F':
            gerar_fraseologia()
        elif escolha.isdigit() and 1 <= int(escolha) <= len(nomes_abas):
            nome_aba_selecionada = nomes_abas[int(escolha) - 1]
            print(f"\n{Colors.BOLD}{Colors.LIGHT_CYAN}--- STATUS DO SISTEMA ---{Colors.ENDC}")
            print(
                f"  {Colors.OKCYAN}► {Colors.BOLD}{nome_aba_selecionada}{Colors.ENDC} {Colors.LIGHT_GREEN}Setor ativado. Status: {Colors.BOLD}[ONLINE]{Colors.ENDC}")
            print(f"  {Colors.OKCYAN}► {Colors.ENDC}{Colors.LIGHT_GREEN}Aguardando consulta...{Colors.ENDC}")

            termos_voltar = ['V', 'VOLTAR']

            while True:
                termo = input(
                    f"\n{Colors.OKBLUE}Termo de busca (digite '{' ou '.join(termos_voltar)}' para retornar ao menu):{Colors.ENDC}\n{Colors.LIGHT_GREEN}➜ {Colors.ENDC}").strip()
                if termo.upper() in termos_voltar:
                    break
                if termo:
                    inicio = time.time()
                    resultados = buscador.buscar_avancada(nome_aba_selecionada, termo)
                    tempo = time.time() - inicio
                    buscador.exibir_resultados_avancados(resultados, termo, nome_aba_selecionada, tempo)
        else:
            print(f"{Colors.RED}✗ Erro de sintaxe: Comando não reconhecido. Tente novamente.{Colors.ENDC}")


if __name__ == "__main__":
    main()