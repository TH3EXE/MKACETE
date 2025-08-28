# MKACETE - Sistema de Busca Avançada em Excel

## 📌 Visão Geral
O **MKACETE** é um sistema de busca avançada desenvolvido para facilitar a pesquisa em planilhas Excel.  
Ele permite encontrar informações rapidamente utilizando diferentes métodos de busca (exata, por relevância e fuzzy/similaridade), com suporte a filtros, cache inteligente e estatísticas de uso.  

O sistema foi projetado para ser simples de usar, mesmo para pessoas sem experiência técnica, bastando ter o arquivo Excel configurado e o Python instalado.

---

## ⚙️ Requisitos

Antes de usar o sistema, é necessário instalar o **Python 3.8+**.  
Em seguida, instale as bibliotecas necessárias executando no terminal:

```bash
pip install pandas openpyxl unidecode
```

---

## 🚀 Como Executar

1. Certifique-se de que o arquivo **`BATMAN.xlsx`** (base de dados) está disponível:
   - O sistema procura primeiro no mesmo diretório do programa.
   - Caso não encontre, tenta buscar no caminho configurado no código:
     ```
     C:\Users\Administrador\Documents\BATMAN.xlsx
     ```
   - Se for usar em outro computador, basta colocar o arquivo no mesmo diretório do script ou ajustar o caminho no código (`main()`).

2. Execute o sistema:
   ```bash
   python mkacete.py
   ```

---

## 📖 Funcionamento Básico

Ao abrir o sistema, será exibido um **menu principal** com as abas disponíveis no Excel.

### Opções principais:
- Digitar o número correspondente a uma aba → abre a pesquisa nessa aba.
- **S** → Mostrar estatísticas do sistema (quantidade de buscas, tempo médio, etc.).
- **C** → Limpar o cache (apagar resultados armazenados em memória).
- **CFG** → Salvar a configuração atual no arquivo `config.json`.
- **0** → Sair do sistema.

### Pesquisando
- Digite o termo que deseja buscar (exemplo: `hospital` ou `ressonancia`).
- Os resultados serão exibidos com os campos disponíveis.
- Para voltar ao menu principal, digite **V**.

---

## 🔍 Funcionalidades

- **Busca Exata**: encontra resultados que contenham o termo digitado.
- **Busca por Relevância**: prioriza colunas mais importantes (ex.: PRESTADOR, PROCEDIMENTO).
- **Busca Fuzzy (Similaridade)**: localiza resultados mesmo com erros de digitação ou grafias diferentes.
- **Filtros Opcionais**: permite restringir a busca a determinadas colunas.
- **Cache Inteligente**: reaproveita resultados de buscas repetidas.
- **Estatísticas**: tempo médio de resposta, total de resultados e número de buscas.
- **Configuração Personalizável** (`config.json`):  
  - Limite de resultados exibidos  
  - Nível de similaridade fuzzy  
  - Ativar/desativar cache  
  - Pesos para colunas  

---

## 🔧 Configuração do Local do Arquivo Excel

Por padrão, o sistema busca o arquivo em duas localizações:
1. Diretório local onde está o script.
2. Caminho fixo configurado no código:
   ```python
   caminho_completo = r'C:\Users\Administrador\Documents\BATMAN.xlsx'
   ```

Se trocar de máquina ou de diretório:
- Coloque o arquivo `BATMAN.xlsx` na mesma pasta do programa **OU**
- Edite a variável `caminho_completo` no código para o novo caminho do arquivo.

---

## 📝 Exemplo de Uso

### Passo 1: Selecionar aba
```
--- Menu Principal ---
[01] PRESTADORES
[02] PROCEDIMENTOS
[S] Estatísticas do sistema
[C] Limpar cache
[CFG] Salvar configuração
[0] Sair
Digite sua escolha: 1
```

### Passo 2: Buscar termo
```
Digite o termo para buscar (ou 'V' para voltar): cardiologia
```

### Resultado exibido:
```
--- ITEM 1 ---
PRESTADOR: hospital central
PROCEDIMENTO: consulta cardiologia
TUSS: 123456
```

---

## 📞 Suporte

Em caso de dúvidas ou sugestões, consulte este documento ou entre em contato diretamente com o desenvolvedor responsável: **@TH3EXE**

---

⚡ **MKACETE - Desenvolvido para tornar a busca em planilhas mais rápida, prática e confiável.**
