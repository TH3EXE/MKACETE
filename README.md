# MKACETE - Sistema de Busca Avançada

## Visão Geral
O MKACETE é um sistema de busca otimizado para análise de planilhas em Excel. 
Ele foi projetado para agilizar pesquisas complexas em grandes volumes de dados, oferecendo algoritmos de busca exata, por relevância e fuzzy (similaridade).

O sistema também conta com geração automática de fraseologias, estatísticas de uso e opções de personalização via arquivo de configuração.

---

## Instalação

### Dependências
Antes de iniciar, certifique-se de ter instalado:

- **Python 3.7 ou superior**: https://www.python.org/downloads/
- **Git**: https://git-scm.com/downloads/win

Durante a instalação do Python, marque a opção **"Add Python to PATH"**.

### Passos de Configuração
1. Baixe o código do projeto em: https://github.com/TH3EXE/MKACETE  
   Ou utilize o Git:
   ```bash
   git clone https://github.com/TH3EXE/MKACETE.git
   ```

2. Extraia os arquivos (se baixado em formato ZIP).

3. Coloque a planilha `BATMAN.xlsx` na mesma pasta do projeto.

4. Execute o instalador automático:
   ```bash
   instalar.bat
   ```

Esse script instala todas as bibliotecas necessárias e prepara o ambiente.

Dependências Python utilizadas:
- `pandas`
- `openpyxl`
- `unidecode`
- `colorama`

---

## Uso

### Executando o sistema
No terminal, navegue até a pasta do projeto e execute:
```bash
python pesquisador.py
```

### Menu Principal
- **Números 1-N**: Seleciona a aba da planilha para pesquisa.  
- **S**: Exibe estatísticas do sistema.  
- **C**: Limpa o cache.  
- **CFG**: Salva configurações atuais.  
- **U**: Atualiza o sistema.  
- **0**: Sai do programa.

---

## Algoritmos de Busca

### Busca Exata
- Localiza correspondências idênticas ao termo pesquisado.

### Busca por Relevância
- Utiliza pesos por coluna e dá prioridade para termos encontrados no início da string ou como palavras completas.

### Busca Fuzzy
- Baseada em similaridade de strings (`SequenceMatcher`).
- Permite configurar o nível de sensibilidade no `config.json`.

---

## Fraseologias Automáticas
O sistema pode gerar frases pré-formatadas a partir dos resultados da busca.  
São suportados dois tipos principais:
- **Positivas**  
- **Negativas**

Exemplo de uso no menu:
```
Digite o termo para buscar: hospital
Deseja gerar fraseologia automática? (S/N):
```

---

## Estatísticas
O sistema coleta as seguintes métricas:
- Total de buscas realizadas.  
- Tempo médio de execução.  
- Total de resultados encontrados.  
- Itens armazenados em cache.  
- Status das configurações.

---

## Personalização

As configurações são feitas no arquivo `config.json`:
```json
{
  "max_resultados": 100,
  "limiar_similaridade": 0.7,
  "habilitar_cache": true
}
```

---

## Exemplos de Uso

### Busca simples
```
Digite o termo para buscar: cardiologia
```

### Busca com filtros
```
Termo: ressonância
Filtros:
  PRESTADOR: clinica
  ZONA: norte
```

### Busca avançada
```
Termo: hospital
Máximo de resultados: 25
```

---

## Solução de Problemas

### Erro de cores no terminal
- Windows: habilitado nativamente.  
- Linux/Mac: compatível sem ajustes.

### Arquivo não encontrado
- Verifique se `BATMAN.xlsx` está no diretório do sistema.  
- Confirme permissões de leitura do arquivo.

### Desempenho lento
- Habilite o cache no `config.json`.  
- Ajuste `max_resultados` para reduzir a carga.  
- Limpe o cache regularmente.

---

## Histórico de Versões

- **v1.0**: Sistema básico de busca.  
- **v2.0**: Implementação de algoritmos múltiplos.  
- **v2.1**: Interface e performance aprimoradas.

---

## Autor
Projeto desenvolvido por **@TH3EXE**.
