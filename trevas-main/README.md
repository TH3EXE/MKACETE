# Sistema de Busca Avançada - BATMAN

## 🚀 Melhorias Implementadas

### 1. **Mecanismo de Busca Otimizado**
- **Busca Multi-Algoritmo**: Combina busca exata, por relevância e fuzzy
- **Sistema de Pesos**: Colunas importantes têm prioridade maior na busca
- **Cache Inteligente**: Armazena resultados para buscas repetidas
- **Busca Fuzzy**: Encontra resultados similares usando similaridade de strings

### 2. **Interface Melhorada**
- **Cores no Terminal**: Interface colorida e organizada
- **Ícones Visuais**: Emojis para melhor identificação dos campos
- **Menu Hierárquico**: Navegação mais intuitiva e organizada
- **Estatísticas em Tempo Real**: Monitoramento de performance

### 3. **Funcionalidades Avançadas**
- **Filtros Múltiplos**: Busca com critérios específicos
- **Configuração Flexível**: Arquivo JSON para personalização
- **Tratamento de Dados**: Limpeza automática de dados vazios
- **Performance**: Otimizações para arquivos grandes

### 4. **Organização dos Dados**
- **Preparação Automática**: Dados são limpos e preparados automaticamente
- **Coluna de Busca**: Texto concatenado para busca mais rápida
- **Validação de Dados**: Tratamento de valores NaN e vazios

## 📋 Requisitos

```bash
pip install pandas openpyxl unidecode
```

## 🎯 Como Usar

### 1. **Executar o Sistema**
```bash
python pesquisador.py
```

### 2. **Menu Principal**
- **Números 1-X**: Selecionar aba para pesquisa
- **S**: Ver estatísticas do sistema
- **C**: Limpar cache
- **CFG**: Salvar configuração
- **0**: Sair

### 3. **Tipos de Busca**
- **Busca Simples**: Termo único
- **Busca com Filtros**: Termo + critérios específicos
- **Busca Avançada**: Termo + configurações personalizadas

## ⚙️ Configuração

O arquivo `config.json` permite personalizar:

- **Limite de resultados**
- **Limiar de similaridade fuzzy**
- **Habilitação de cache**
- **Pesos das colunas**
- **Configurações de interface**

## 🔍 Algoritmos de Busca

### 1. **Busca Exata**
- Encontra correspondências exatas do termo
- Maior prioridade nos resultados

### 2. **Busca por Relevância**
- Usa sistema de pesos por coluna
- Bônus para início de string e palavras completas

### 3. **Busca Fuzzy**
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
- **Ícones**: Emojis para identificação rápida
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
