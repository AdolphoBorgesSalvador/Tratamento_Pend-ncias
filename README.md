# Projeto de Consolidação de Pendências

Este projeto utiliza Python e a biblioteca Pandas para consolidar dados de diferentes planilhas relacionadas a pendências e outras métricas operacionais. O objetivo é realizar junções de dados a partir de múltiplos arquivos Excel, construir tabelas dinâmicas e gerar uma planilha consolidada contendo informações detalhadas sobre materiais, estoque e pendências.

## Pré-requisitos

- Python 3.7 ou superior
- Biblioteca Pandas

Para instalar a biblioteca Pandas, execute:

```bash
pip install pandas
```

## Estrutura dos Arquivos

Os arquivos utilizados no projeto estão localizados no diretório `C:\Users\fsp_adolpho.salvador\Desktop\Konica Minolta\Desktop Cloud - Documentos\Desktop\Teste Pendencias`. Os principais arquivos são:

- **PEND.XLSX**: Dados de pendências iniciais.
- **pend geral.XLSX**: Dados de pendências gerais.
- **zstok.XLSX**: Informações de estoque por centro.
- **fup.XLSX**: Informações de pedidos e datas.
- **ZTMM069.XLSX** e **ZTMM085.XLSX**: Informações de peças.
- **atuais.xlsx**: Código atual de materiais.
- **Classificação das Peças.xlsx**: Classificação e durabilidade das peças.

## Funcionalidades

### 1. Leitura e Seleção de Colunas
Leitura dos dados de pendências e exibição de colunas específicas para análise inicial.

### 2. Tabelas Dinâmicas
Criação de várias tabelas dinâmicas a partir dos dados importados:
- **PEND.XLSX**: Tabelas de quantidades pendentes e informações de material.
- **pend geral.XLSX**: Tabela consolidada de pendências gerais.
- **zstok.XLSX**: Tabelas de estoque disponível e controle de qualidade.
- **fup.XLSX**: Tabela de pedidos por mês e ano.
- **ZTMM069.XLSX** e **ZTMM085.XLSX**: Tabelas de montante e modelos de peças.

### 3. Junção e Consolidação dos Dados
Os dados de todas as tabelas dinâmicas são integrados em uma única tabela final:
- Junção das pendências com estoque, pedidos e informações adicionais.
- Inclusão dos dados de códigos atuais e classificações de peças.
- Preenchimento de valores nulos com zero para garantir consistência.

### 4. Exportação dos Dados Consolidados
Exportação da tabela consolidada final para um arquivo Excel (`MAPA Pendencia.xlsx`) no diretório especificado.

## Como Executar o Script

1. Certifique-se de que todos os arquivos necessários estão no diretório correto.
2. Execute o script `consolida_pendencias.py` (ou o nome que preferir dar ao seu script) no terminal:

```bash
python consolida_pendencias.py
```

3. O script irá gerar um arquivo consolidado `MAPA Pendencia.xlsx` com todas as informações integradas.

## Notas

- Certifique-se de que as colunas dos arquivos de entrada correspondem aos nomes utilizados no código.
- Ajuste os caminhos dos arquivos se necessário.
  
Este projeto foi desenvolvido para facilitar a análise de pendências e ajudar na gestão de estoque e planejamento de materiais.
