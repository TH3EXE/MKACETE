import pandas as pd
import os

# Altere o nome do seu arquivo Excel aqui
NOME_ARQUIVO = 'BATMAN🦇.xlsx'


def carregar_planilha():
    """
    Carrega a planilha Excel e retorna um dicionário, onde cada
    chave é o nome de uma aba e o valor é o DataFrame correspondente.
    """
    try:
        print("Carregando dados da planilha...")
        # Usa sheet_name=None para ler todas as abas
        dados_abas = pd.read_excel(NOME_ARQUIVO, sheet_name=None)
        print("Planilha carregada com sucesso.")
        return dados_abas
    except FileNotFoundError:
        print(f"Erro: O arquivo '{NOME_ARQUIVO}' não foi encontrado.")
        print("Certifique-se de que o nome está correto e o arquivo está no mesmo diretório do script.")
        return None
    except Exception as e:
        print(f"Ocorreu um erro ao carregar o arquivo Excel: {e}")
        return None


def buscar_dados(df, termo):
    """
    Busca um termo em todas as colunas de um DataFrame,
    ignorando a diferença entre maiúsculas e minúsculas.
    """
    # Converte todo o DataFrame para string para facilitar a busca
    df_str = df.astype(str)

    # Cria uma máscara para encontrar o termo em qualquer coluna
    # `case=False` ignora maiúsculas/minúsculas
    mask = df_str.apply(lambda row: row.str.contains(termo, case=False, na=False)).any(axis=1)

    return df[mask]


def exibir_resultados(resultados, termo):
    """
    Imprime os resultados da busca no terminal.
    """
    if resultados.empty:
        print(f"\nNenhum resultado encontrado para o termo '{termo}'.")
    else:
        print(f"\n--- Resultados encontrados para '{termo}' ---")
        # Configura o pandas para exibir todas as colunas
        pd.set_option('display.max_columns', None)
        print(resultados)
        print("------------------------------------------")
    print()


def menu_principal():
    """
    Função principal que gerencia a navegação e a busca.
    """
    dados_abas = carregar_planilha()
    if not dados_abas:
        return

    while True:
        print("\n### Menu de Navegação ###")
        print("Escolha uma aba para pesquisar:")

        nomes_abas = list(dados_abas.keys())
        for i, nome in enumerate(nomes_abas, 1):
            print(f"  [{i}] {nome}")
        print("  [0] Sair")

        try:
            escolha = int(input("\nDigite o número da aba (ou 0 para sair): "))

            if escolha == 0:
                print("Saindo do programa. Até mais!")
                break

            if 1 <= escolha <= len(nomes_abas):
                nome_aba_selecionada = nomes_abas[escolha - 1]
                df_selecionado = dados_abas[nome_aba_selecionada]

                print(f"\nVocê está na aba: '{nome_aba_selecionada}'.")

                while True:
                    termo_busca = input("Digite o termo para buscar (ou 'voltar' para o menu principal): ").strip()

                    if termo_busca.lower() == 'voltar':
                        break

                    if not termo_busca:
                        print("O termo de busca não pode ser vazio.")
                        continue

                    resultados = buscar_dados(df_selecionado, termo_busca)
                    exibir_resultados(resultados, termo_busca)

            else:
                print("Escolha inválida. Por favor, digite um número da lista.")

        except ValueError:
            print("Entrada inválida. Por favor, digite um número.")


if __name__ == "__main__":
    menu_principal()