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

        # --- NOVA LÓGICA DE EXIBIÇÃO POR CATEGORIA ---
        if nome_aba.lower() == 'infiltracao':
            self._exibir_resultados_por_categoria(resultados)
        else:
            self._exibir_generico_futurista(resultados)

        print(f"\n{Colors.LIGHT_CYAN}-----------------------------------------{Colors.ENDC}")

    def _exibir_resultados_por_categoria(self, resultados):
        """
        Organiza e exibe os resultados da aba 'infiltracao' em blocos separados.
        """
        # Normaliza o termo de busca para a categorização
        resultados['_TEXTO_BUSCA'] = resultados['_TEXTO_BUSCA'].apply(lambda x: unidecode(str(x)).lower())

        # Separa os resultados em categorias
        mask_infiltracao = resultados['_TEXTO_BUSCA'].str.contains('infiltracao|bloqueio|infiltração')
        mask_retirada = resultados['_TEXTO_BUSCA'].str.contains('retirada|remocao|remover')

        # Cria dataframes para cada categoria
        df_infiltracao = resultados[mask_infiltracao]
        df_retirada = resultados[mask_retirada]

        # Filtra os resultados que não se encaixam nas categorias acima
        df_outros = resultados[~(mask_infiltracao | mask_retirada)]

        # Exibe os blocos
        self._exibir_bloco_categoria("Resultados para INFILTRAÇÃO", df_infiltracao)
        self._exibir_bloco_categoria("Resultados para RETIRADA DE MEDICAÇÃO", df_retirada)

        if not df_outros.empty:
            self._exibir_bloco_categoria("Outros Resultados", df_outros)

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
            print(f"\n{Colors.BOLD}{Colors.LIGHT_CYAN}═══════════════ [{i:03d}] Registro de Dados ═══════════════{Colors.ENDC}")

            # Padrões de cores para os campos
            key_color = Colors.LIGHT_BLUE
            value_color = Colors.OKGREEN

            for col, val in row.items():
                if str(val).strip() and col != '_TEXTO_BUSCA':
                    print(f"  {key_color}{str(col).upper()}:{Colors.ENDC} {value_color}{val}{Colors.ENDC}")

            print(f"{Colors.BOLD}{Colors.LIGHT_CYAN}═══════════════════════════════════════════════════════{Colors.ENDC}")


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


def main():
    """Função principal que executa a interface de linha de comando."""
    colorama.init()

    # Get terminal size
    try:
        terminal_width = os.get_terminal_size().columns
    except OSError:
        terminal_width = 80  # default if terminal size is not available

    # Helper function to center text
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
        print(center_text(f"{Colors.YELLOW}Selecione um setor de dados para iniciar a busca:{Colors.ENDC}",
                          terminal_width + 12))

        # Novo layout em colunas
        colunas = 2
        total_abas = len(nomes_abas)
        abas_por_coluna = (total_abas + colunas - 1) // colunas

        # Calculate a larger padding to help with centering
        max_aba_length = max(len(aba) for aba in nomes_abas)
        col_width = max_aba_length + 8  # Extra padding for number and symbol
        menu_width = colunas * col_width

        # Ensure a minimal width for alignment, even with a small terminal
        min_col_width = 30
        if col_width < min_col_width:
            col_width = min_col_width

        # Calculate the proper left margin for centering the entire block
        total_menu_width = colunas * (col_width + 2) # +2 for "  " prefix
        left_margin = max(0, (terminal_width - total_menu_width) // 2)

        for i in range(abas_por_coluna):
            linha = " " * left_margin
            for j in range(colunas):
                index = i + j * abas_por_coluna
                if index < total_abas:
                    nome_aba = nomes_abas[index]
                    linha += f"  {Colors.OKCYAN}► [{index + 1:02d}]{Colors.ENDC} {nome_aba:<{col_width - 8}}"
                else:
                    # Pad with spaces to keep the columns aligned
                    linha += " " * (col_width + 2)

            # Only print the line if it contains content
            if linha.strip():
                print(linha)


        print(center_text(f"\n  {Colors.OKCYAN}► [S]{Colors.ENDC} Relatório de status", terminal_width + 12))
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
        elif escolha.isdigit() and 1 <= int(escolha) <= len(nomes_abas):
            nome_aba_selecionada = nomes_abas[int(escolha) - 1]
            print(f"\n{Colors.BOLD}{Colors.LIGHT_CYAN}--- STATUS DO SISTEMA ---{Colors.ENDC}")
            print(f"  {Colors.OKCYAN}► {Colors.BOLD}{nome_aba_selecionada}{Colors.ENDC} {Colors.LIGHT_GREEN}Setor ativado. Status: {Colors.BOLD}[ONLINE]{Colors.ENDC}")
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