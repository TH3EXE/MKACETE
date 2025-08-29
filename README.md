# MKACETE - Sistema de Busca Avançada

## 🚀 Manual Rápido do MKACETE
Manual Rápido do MKACETE
Este manual rápido vai ajudar você a colocar o sistema para funcionar.

### 1. **O que Instalar**

Você só precisa de dois programas no seu computador. Clique nos links para baixar:

- **Python**: https://www.python.org/downloads/
- Importante: Durante a instalação, marque a caixa "Add Python to PATH".
  
- **Git**: https://git-scm.com/downloads/win

### 2. **Configurar o Projeto**
- **Baixe o código do projeto MKACETE**: clicando aqui: https://github.com/TH3EXE/MKACETE.
- Extraia o arquivo ZIP e coloque a sua planilha BATMAN.xlsx na mesma pasta.

### 3. **Rodar o Sistema**
- **Abra o Prompt de Comando do Windows.**
- **Use o comando cd para ir até a pasta onde estão os arquivos. Por exemplo:**: cd C:\Users\SeuUsuario\Downloads\MKACETE
- **Instale as bibliotecas que o sistema precisa com este comando:**: py -m pip install pandas openpyxl unidecode colorama
- **Para iniciar o sistema, digite:**: python pesquisador.py

### 4. **Organização dos Dados**
- **Preparação Automática**: Dados são limpos e preparados automaticamente
- **Coluna de Busca**: Texto concatenado para busca mais rápida
- **Validação de Dados**: Tratamento de valores NaN e vazios

- Pronto! O sistema vai abrir e você pode começar a usar.


### 1.1. **Menu Principal**
- **Números 1-X**: Selecionar aba para pesquisa
- **S**: Ver estatísticas do sistema
- **C**: Limpar cache
- **CFG**: Salvar configuração
- **U**: Atualizar o Sistema
- **0**: Sair

## 🔍 Algoritmos de Busca

### 1.2. **Busca Exata**
- Encontra correspondências exatas do termo
- Maior prioridade nos resultados

### 1.3. **Busca por Relevância**
- Usa sistema de pesos por coluna
- Bônus para início de string e palavras completas

### 1.4. **Busca Fuzzy**
- Similaridade de strings usando SequenceMatcher
- Configurável via limiar de similaridade

## 📊 Estatísticas

O sistema monitora:
- Total de buscas realizadas
- Tempo médio de busca
- Total de resultados encontrados
- Itens em cache
- Status da configuração

## 🎨 Interface Visual

- **Cores**: Diferentes cores para tipos de informação
- **Formatação**: Resultados organizados e legíveis
- **Navegação**: Menu intuitivo e responsivo

## 🚀 Performance

- **Cache**: Resultados armazenados para buscas repetidas
- **Otimização de Dados**: Preparação automática dos dados
- **Busca Eficiente**: Múltiplos algoritmos otimizados
- **Gerenciamento de Memória**: Cache com limite configurável

## 🔧 Personalização

### Editar Configuração
```json
{
  "max_resultados": 100,
  "limiar_similaridade": 0.7,
  "habilitar_cache": true
}
```

### Adicionar Novas Abas
O sistema detecta automaticamente todas as abas do Excel e aplica formatação apropriada.

## 📝 Exemplos de Uso

### Busca Simples
```
Digite o termo para buscar: hospital
```

### Busca com Filtros
```
Termo: cardiologia
Filtros:
  PRESTADOR: centro
  ZONA: sul
```

### Busca Avançada
```
Termo: ressonância
Máximo de resultados: 25
```

## 🐛 Solução de Problemas

### Erro de Cores no Terminal
- Windows: Suporte nativo para cores ANSI
- Linux/Mac: Funciona normalmente

### Arquivo Não Encontrado
- Verificar se `BATMAN.xlsx` está no mesmo diretório
- Verificar permissões de leitura

### Performance Lenta
- Verificar se cache está habilitado
- Ajustar `max_resultados` na configuração
- Limpar cache se necessário

## 🔄 Versões

- **v1.0**: Sistema básico de busca
- **v2.0**: Sistema avançado com múltiplos algoritmos
- **v2.1**: Interface melhorada e otimizações de performance

## 📞 Suporte

Para dúvidas ou sugestões, consulte a documentação ou entre em contato com o desenvolvedor responsável: @TH3EXE.

---


**Desenvolvido com a força da preguiça para otimizar a busca em dados médicos** 


