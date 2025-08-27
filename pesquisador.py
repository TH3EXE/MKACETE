import pandas as pd
import os
from unidecode import unidecode
from glob import glob


class MecanismoBusca:
    """
    Classe principal para gerenciar o carregamento de dados e a busca.
    """

    def __init__(self, nome_arquivo_excel):
        """
        Inicializa o buscador com o nome do arquivo Excel.

        Args:
            nome_arquivo_excel (str): O nome do arquivo .xlsx.
        """
        self.nome_arquivo_excel = nome_arquivo_excel
        self.dados_abas = self._carregar_dados()

    def _carregar_dados(self):
        """
        Carrega todas as abas (sheets) do arquivo Excel especificado.
        Cada aba é tratada como um "banco de dados" para busca.
        """
        print("Carregando o banco de dados...")

        # Verifica se o arquivo existe no mesmo diretório do script
        if not os.path.exists(self.nome_arquivo_excel):
            print(f"Erro: Arquivo não encontrado: '{self.nome_arquivo_excel}'")
            print("Certifique-se de que o arquivo está no mesmo diretório do script.")
            return None

        try:
            # Carrega o arquivo Excel e lê todas as abas
            return pd.read_excel(self.nome_arquivo_excel, sheet_name=None)
        except Exception as e:
            print(f"Ocorreu um erro ao carregar o arquivo Excel: {e}")
            return None

    def buscar_na_aba(self, nome_aba, termo):
        """
        Busca um termo em uma aba específica. A busca não diferencia
        maiúsculas/minúsculas e ignora acentos.

        Args:
            nome_aba (str): O nome da aba onde a busca será realizada.
            termo (str): O termo a ser pesquisado.

        Returns:
            DataFrame: O DataFrame com os resultados encontrados.
        """
        if not self.dados_abas or nome_aba not in self.dados_abas:
            return pd.DataFrame()

        df = self.dados_abas[nome_aba]
        termo_limpo = unidecode(termo).lower().strip()

        # Linha corrigida para remover o FutureWarning
        df_preparado = df.astype(str).map(lambda x: unidecode(x).lower())

        # Cria uma máscara para encontrar o termo em qualquer coluna da aba
        mascara = df_preparado.apply(lambda row: row.str.contains(termo_limpo, na=False)).any(axis=1)

        return df[mascara]

    def exibir_resultados(self, resultados, termo, nome_aba):
        """
        Exibe os resultados da busca de forma organizada e limpa,
        tratando os valores vazios e N/A.

        Args:
            resultados (DataFrame): O DataFrame com os resultados da busca.
            termo (str): O termo buscado.
            nome_aba (str): O nome da aba onde a busca foi feita.
        """
        if resultados.empty:
            print(f"\nNenhum resultado encontrado para '{termo}' na aba '{nome_aba}'.")
            return

        print(f"\n--- Resultados encontrados para '{termo}' na aba '{nome_aba}' ---")

        # Mapeamento para funções de exibição personalizadas por aba
        exibidores = {
            'INFILTRAÇÃO': self._exibir_prestador,
            'CENTROS CLINICOS | R.P': self._exibir_prestador,
            'HOSPITAIS | R.P': self._exibir_prestador,
            'QUALIVIDA | R.P': self._exibir_prestador,
            'NUCLEOS DE TERAPIAS | R.P': self._exibir_prestador,
            'NOTRELABS | R.P': self._exibir_prestador,
            'CDB | R.C': self._exibir_prestador,
            'HERMES PARDINI | R.C': self._exibir_prestador,
            'LAVOISIER - DELBONI - | R.C': self._exibir_prestador,
            'PROCEDIMENTOS': self._exibir_procedimento,
            'PROCEDIMENTOS (EXTRA FUNIL)': self._exibir_procedimento_extra_funil,
            'DIU': self._exibir_diu
        }

        # Chama o exibidor específico ou o genérico se não houver um
        exibidor = exibidores.get(nome_aba, self._exibir_generico)
        exibidor(resultados)

        print("------------------------------------------")

    def _exibir_prestador(self, df):
        for _, row in df.iterrows():
            if pd.notna(row.get('PRESTADOR')):
                print(f"**PRESTADOR**: {row['PRESTADOR']}")

            zona_regiao = row.get('ZONA', row.get('REGIÃO', None))
            if pd.notna(zona_regiao) and str(zona_regiao).strip() != '':
                print(f"  ZONA OU REGIÃO: {zona_regiao}")

            for col in ['CNPJ', 'CD PESSOA', 'RAZAO SOCIAL', 'ENDEREÇO']:
                valor = row.get(col)
                if pd.notna(valor) and str(valor).strip() != '':
                    print(f"  {col.upper()}: {valor}")
            print("---")

    def _exibir_procedimento(self, df):
        for _, row in df.iterrows():
            if pd.notna(row.get('PROCEDIMENTOS')):
                print(f"**PROCEDIMENTO**: {row['PROCEDIMENTOS']}")

            for col in ['TUSS', 'AMB', 'CBHPM', 'OBSERVACAO']:
                valor = row.get(col)
                if pd.notna(valor) and str(valor).strip() != '':
                    print(f"  {col.upper()}: {valor}")
            print("---")

    def _exibir_procedimento_extra_funil(self, df):
        for _, row in df.iterrows():
            proc_tuss = row.get('PROCEDIMENTOS TUSS')
            tuss = row.get('TUSS')

            if pd.notna(proc_tuss):
                print(f"  PROCEDIMENTO (TUSS): {proc_tuss}")
                if pd.notna(tuss):
                    print(f"  CÓDIGO TUSS: {tuss}")

            proc_amb = row.get('PROCEDIMENTOS AMB')
            amb = row.get('AMB')
            if pd.notna(proc_amb):
                print(f"  PROCEDIMENTO (AMB): {proc_amb}")
                if pd.notna(amb):
                    print(f"  CÓDIGO AMB: {amb}")
            print("---")

    def _exibir_diu(self, df):
        for _, row in df.iterrows():
            procedimento = row.get('PROCEDIMENTO')
            if pd.notna(procedimento):
                print(f"  **PROCEDIMENTO**: {procedimento}")

            for col in ['TUSS', 'AMB', 'OBSERVAÇÃO', 'OBSERVAÇÃO 02']:
                valor = row.get(col)
                if pd.notna(valor) and str(valor).strip() != '':
                    print(f"  **{col.upper()}**: {valor}")
            print("---")

    def _exibir_generico(self, df):
        """Exibe todas as colunas de um DataFrame, ignorando valores vazios."""
        for i, row in df.iterrows():
            print(f"**Item {i+1}**:")
            for coluna, valor in row.items():
                if pd.notna(valor) and str(valor).strip() != '':
                    print(f"  **{str(coluna).upper()}**: {valor}")
            print("---")


def main():
    """
    Função principal que inicia e gerencia o loop de navegação e busca.
    """
    # ATENÇÃO: Altere o nome do arquivo se necessário
    nome_arquivo = 'BATMAN.xlsx'

    buscador = MecanismoBusca(nome_arquivo)

    if not buscador.dados_abas:
        return

    nomes_abas = list(buscador.dados_abas.keys())

    while True:
        print("\n--- Menu de Navegação ---")
        print("Escolha uma aba para pesquisar:")

        for i, nome in enumerate(nomes_abas, 1):
            print(f"  [{i}] {nome}")
        print("  [0] Sair")

        try:
            escolha = input("\nDigite o número da aba (ou 0 para sair): ").strip()

            if not escolha.isdigit():
                print("Escolha inválida. Por favor, digite um número.")
                continue

            escolha = int(escolha)

            if escolha == 0:
                print("Saindo do programa. Até mais!")
                break

            if 1 <= escolha <= len(nomes_abas):
                nome_aba_selecionada = nomes_abas[escolha - 1]
                print(f"\nVocê está na aba: '{nome_aba_selecionada}'.")

                while True:
                    termo_busca = input("Digite o termo para buscar (ou 'voltar' para o menu principal): ").strip()
                    if termo_busca.lower() == 'voltar':
                        break

                    if not termo_busca:
                        print("Por favor, digite um termo de busca válido.")
                        continue

                    resultados = buscador.buscar_na_aba(nome_aba_selecionada, termo_busca)
                    buscador.exibir_resultados(resultados, termo_busca, nome_aba_selecionada)
            else:
                print("Escolha inválida. Por favor, digite um número da lista.")

        except ValueError:
            print("Entrada inválida. Por favor, digite um número.")


if __name__ == "__main__":
    main()
