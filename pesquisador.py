import pandas as pd
import os
import time
from unidecode import unidecode
from difflib import SequenceMatcher
import json
import warnings
import colorama
import subprocess

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
        print(f"{Colors.HEADER}Carregando o banco de dados...{Colors.ENDC}")

        if not os.path.exists(self.nome_arquivo_excel):
            print(f"{Colors.FAIL}Erro: Arquivo não encontrado: '{self.nome_arquivo_excel}'{Colors.ENDC}")
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
            print(f"{Colors.OKGREEN}✓ Banco de dados carregado em {tempo_carregamento:.2f}s{Colors.ENDC}")
            print(f"{Colors.OKBLUE}Total de abas: {len(self.dados_abas)}{Colors.ENDC}")

            self._mostrar_estatisticas_abas()

        except Exception as e:
            print(f"{Colors.FAIL}Erro ao carregar arquivo Excel: {e}{Colors.ENDC}")

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
        print(f"\n{Colors.UNDERLINE}Estatísticas das Abas:{Colors.ENDC}")
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
            print(f"\n{Colors.WARNING}Nenhum resultado encontrado para '{termo}' na aba '{nome_aba}'.{Colors.ENDC}")
            return

        print(f"\n{Colors.HEADER}--- Resultados da Busca ---{Colors.ENDC}")
        print(f"{Colors.OKBLUE}Termo buscado: {Colors.BOLD}{termo}{Colors.ENDC}")
        print(f"{Colors.OKBLUE}Aba: {Colors.BOLD}{nome_aba}{Colors.ENDC}")
        print(f"{Colors.OKBLUE}Total de resultados: {Colors.BOLD}{len(resultados)}{Colors.ENDC}")
        if tempo_busca:
            print(f"{Colors.OKBLUE}Tempo de busca: {Colors.BOLD}{tempo_busca:.3f}s{Colors.ENDC}")
        print(f"{Colors.HEADER}------------------------------------------{Colors.ENDC}")

        # --- NOVA LÓGICA DE EXIBIÇÃO POR CATEGORIA ---
        if nome_aba.lower() == 'infiltracao':
            self._exibir_resultados_por_categoria(resultados)
        else:
            self._exibir_generico(resultados)

        print(f"{Colors.HEADER}------------------------------------------{Colors.ENDC}")

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
            print(f"\n{Colors.BOLD}{Colors.OKBLUE}----- {titulo} ({len(df)}) -----{Colors.ENDC}")
            self._exibir_generico(df)

    def _exibir_generico(self, df):
        """Exibe todos os dados de forma genérica e limpa, sem ocultar informações."""
        for i, (_, row) in enumerate(df.iterrows(), 1):
            print(f"\n{Colors.OKGREEN}--- ITEM {i} ---{Colors.ENDC}")
            for col, val in row.items():
                if str(val).strip() and col != '_TEXTO_BUSCA':
                    print(f"  {Colors.BOLD}{str(col).upper()}:{Colors.ENDC} {val}")

    def mostrar_estatisticas(self):
        """Mostra estatísticas de uso do sistema."""
        print(f"\n{Colors.HEADER}--- Estatísticas do Sistema ---{Colors.ENDC}")
        print(
            f"{Colors.OKBLUE}Total de buscas realizadas: {Colors.BOLD}{self.estatisticas['total_buscas']}{Colors.ENDC}")
        print(
            f"{Colors.OKBLUE}Tempo médio de busca: {Colors.BOLD}{self.estatisticas['tempo_medio_busca']:.3f}s{Colors.ENDC}")
        print(
            f"{Colors.OKBLUE}Total de resultados encontrados: {Colors.BOLD}{self.estatisticas['resultados_encontrados']}{Colors.ENDC}")
        print(f"{Colors.OKBLUE}Itens em cache: {Colors.BOLD}{len(self.cache_busca)}{Colors.ENDC}")
        print(
            f"{Colors.OKBLUE}Configuração ativa: {Colors.BOLD}{'Sim' if self.config['habilitar_cache'] else 'Não'}{Colors.ENDC}")
        print(f"{Colors.HEADER}------------------------------------------{Colors.ENDC}")

    def limpar_cache(self):
        """Limpa o cache de busca."""
        self.cache_busca.clear()
        print(f"{Colors.OKGREEN}✓ Cache limpo com sucesso!{Colors.ENDC}")

    def salvar_configuracao(self):
        """Salva a configuração atual em arquivo JSON."""
        try:
            with open('config.json', 'w', encoding='utf-8') as f:
                json.dump(self.config, f, indent=2, ensure_ascii=False)
            print(f"{Colors.OKGREEN}✓ Configuração salva em 'config.json'{Colors.ENDC}")
        except Exception as e:
            print(f"{Colors.FAIL}Erro ao salvar configuração: {e}{Colors.ENDC}")


def atualizar_sistema(repo_url):
    """
    Verifica e atualiza o sistema a partir do repositório Git.
    """
    try:
        if not os.path.exists('.git'):
            print(f"{Colors.WARNING}Inicializando o repositório Git local...{Colors.ENDC}")
            subprocess.run(['git', 'init'], check=True)
            subprocess.run(['git', 'remote', 'add', 'origin', repo_url], check=True)
            print(f"{Colors.OKGREEN}✓ Repositório Git local configurado.{Colors.ENDC}")

        print(f"{Colors.OKCYAN}Buscando atualizações no repositório...{Colors.ENDC}")

        # O subprocess.run com check=True levanta uma exceção se o comando falhar
        subprocess.run(['git', 'pull', 'origin', 'main'], check=True, capture_output=True, text=True)

        print(f"\n{Colors.OKGREEN}✓ O sistema foi atualizado com sucesso!{Colors.ENDC}")
        print(f"{Colors.OKGREEN}Por favor, reinicie o programa para aplicar as alterações.{Colors.ENDC}")

    except FileNotFoundError:
        print(f"\n{Colors.FAIL}Erro: O comando 'git' não foi encontrado.{Colors.ENDC}")
        print(f"{Colors.FAIL}Por favor, instale o Git em sua máquina e adicione-o ao PATH.{Colors.ENDC}")
        print(f"{Colors.OKBLUE}Link para download: https://git-scm.com/downloads{Colors.ENDC}")
    except subprocess.CalledProcessError as e:
        print(f"\n{Colors.FAIL}Erro durante a atualização do Git:{Colors.ENDC}")
        print(f"{Colors.FAIL}Detalhes: {e.stderr.strip()}{Colors.ENDC}")
    except Exception as e:
        print(f"\n{Colors.FAIL}Ocorreu um erro inesperado durante a atualização: {e}{Colors.ENDC}")


def main():
    """Função principal que executa a interface de linha de comando."""
    colorama.init()
    repo_url = "https://github.com/TH3EXE/MKACETE.git"

    print(f"\n{Colors.BOLD}{Colors.OKGREEN}")
    print("=" * 60)
    print("      MKACETE: SISTEMA DE BUSCA INTELIGENTE")
    print("=" * 60)
    print(f"{Colors.ENDC}")

    caminho_local = 'BATMAN.xlsx'
    caminho_completo = r'C:\Users\Administrador\Documents\BATMAN.xlsx'

    if os.path.exists(caminho_local):
        nome_arquivo = caminho_local
        print(f"{Colors.OKGREEN}✓ Encontrado: {caminho_local}{Colors.ENDC}")
    elif os.path.exists(caminho_completo):
        nome_arquivo = caminho_completo
        print(f"{Colors.OKGREEN}✓ Encontrado: {caminho_completo}{Colors.ENDC}")
    else:
        print(f"{Colors.FAIL}✗ Erro: Arquivo BATMAN.xlsx não encontrado em nenhum dos caminhos:{Colors.ENDC}")
        print(f"  Local: {caminho_local}")
        print(f"  Completo: {caminho_completo}")
        return

    buscador = MecanismoBuscaAvancado(nome_arquivo)

    if not buscador.dados_abas:
        print(f"{Colors.FAIL}✗ Erro: Não foi possível carregar o banco de dados.{Colors.ENDC}")
        return

    nomes_abas = list(buscador.dados_abas.keys())

    while True:
        print(f"\n{Colors.HEADER}{'=' * 60}{Colors.ENDC}")
        print(f"{Colors.BOLD}MENU PRINCIPAL{Colors.ENDC}")
        print(f"{Colors.OKBLUE}Escolha uma opção ou selecione uma aba para buscar:{Colors.ENDC}")

        for i, nome_aba in enumerate(nomes_abas, 1):
            print(f"  {Colors.OKCYAN}[{i:02d}]{Colors.ENDC} {nome_aba}")

        print(f"\n  {Colors.OKCYAN}[S]{Colors.ENDC} Estatísticas")
        print(f"  {Colors.OKCYAN}[C]{Colors.ENDC} Limpar Cache")
        print(f"  {Colors.OKCYAN}[CFG]{Colors.ENDC} Salvar Configuração")
        print(f"  {Colors.OKCYAN}[U]{Colors.ENDC} Atualizar o Sistema")
        print(f"  {Colors.OKCYAN}[0]{Colors.ENDC} Sair")
        print(f"{Colors.HEADER}{'-' * 60}{Colors.ENDC}")

        escolha = input(f"{Colors.OKGREEN}➜ {Colors.ENDC}").strip().upper()

        if escolha == '0':
            print(f"\n{Colors.OKGREEN}Obrigado por usar o sistema! Até mais!{Colors.ENDC}")
            break
        elif escolha == 'S':
            buscador.mostrar_estatisticas()
        elif escolha == 'C':
            buscador.limpar_cache()
        elif escolha == 'CFG':
            buscador.salvar_configuracao()
        elif escolha == 'U':
            confirmacao = input(
                f"\n{Colors.WARNING}Deseja realmente atualizar o sistema? (S/N): {Colors.ENDC}").strip().upper()
            if confirmacao == 'S':
                atualizar_sistema(repo_url)
            elif confirmacao == 'N':
                print(f"{Colors.OKBLUE}Atualização cancelada. Voltando ao menu...{Colors.ENDC}")
            else:
                print(f"{Colors.FAIL}Resposta inválida. Por favor, digite 'S' para sim ou 'N' para não.{Colors.ENDC}")
        elif escolha.isdigit() and 1 <= int(escolha) <= len(nomes_abas):
            nome_aba_selecionada = nomes_abas[int(escolha) - 1]
            print(f"\n{Colors.OKGREEN}✓ Aba selecionada: '{nome_aba_selecionada}'{Colors.ENDC}")
            termos_voltar = ['V', 'VOLTAR']
            while True:
                termo = input(
                    f"\n{Colors.OKBLUE}Qual termo deseja buscar (digite '{' ou '.join(termos_voltar)}' para voltar)?{Colors.ENDC}\n{Colors.OKGREEN}➜ {Colors.ENDC}").strip()
                if termo.upper() in termos_voltar:
                    break
                if termo:
                    inicio = time.time()
                    resultados = buscador.buscar_avancada(nome_aba_selecionada, termo)
                    tempo = time.time() - inicio
                    buscador.exibir_resultados_avancados(resultados, termo, nome_aba_selecionada, tempo)
        else:
            print(f"{Colors.FAIL}✗ Opção inválida. Tente novamente.{Colors.ENDC}")


if __name__ == "__main__":
    main()